package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"time"
)

// ────────────────────────────────────────────────
// Admin-managed LLM settings
// ────────────────────────────────────────────────

type LLMSettings struct {
	URL string `json:"url"` // e.g. "http://localhost:8080"
}

var (
	llmSettings     LLMSettings
	llmSettingsMu   sync.RWMutex
	llmSettingsFile = filepath.Join(dataDir, "llm_settings.json")
)

func loadLLMSettings() {
	llmSettingsMu.Lock()
	defer llmSettingsMu.Unlock()
	data, err := os.ReadFile(llmSettingsFile)
	if err != nil {
		if !os.IsNotExist(err) {
			log.Printf("Error reading LLM settings: %v", err)
		}
		return
	}
	if err := json.Unmarshal(data, &llmSettings); err != nil {
		log.Printf("Error decoding LLM settings: %v", err)
	}
	log.Printf("LLM settings loaded: URL=%s", llmSettings.URL)
}

func saveLLMSettings() error {
	llmSettingsMu.RLock()
	data, err := json.MarshalIndent(llmSettings, "", "  ")
	llmSettingsMu.RUnlock()
	if err != nil {
		return err
	}
	data = append(data, '\n')
	return os.WriteFile(llmSettingsFile, data, 0644)
}

func GetLLMURL() string {
	llmSettingsMu.RLock()
	defer llmSettingsMu.RUnlock()
	return llmSettings.URL
}

func SetLLMURL(url string) error {
	llmSettingsMu.Lock()
	llmSettings.URL = strings.TrimSpace(url)
	llmSettingsMu.Unlock()
	return saveLLMSettings()
}

// ────────────────────────────────────────────────
// AI prompt reference resolution
// ────────────────────────────────────────────────

// ResolveAIPrompt replaces {{A1}}, {{A1:B3}}, {{project/sheet/A1}}, {{CellName}} etc.
// with their resolved values suitable for natural-language prompt consumption.
//   - Single cell  → plain value
//   - 1D range     → comma-separated values
//   - 2D range     → markdown table
//   - Cell names are first resolved to coordinates
func ResolveAIPrompt(prompt, projectName, sheetName string) string {
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
	if s == nil {
		return prompt
	}

	// --- Pre-pass: resolve cell-name references to coordinate references ---

	// Cross-sheet cell-name: {{project/.../sheet/CellName}} -> {{project/.../sheet/ColRow}}
	crossSheetNamePre := regexp.MustCompile(`\{\{((?:[^/\{\}]+/)+)([^/\{\}]+)/([A-Za-z_]\w*)\}\}`)
	prompt = crossSheetNamePre.ReplaceAllStringFunc(prompt, func(match string) string {
		sub := crossSheetNamePre.FindStringSubmatch(match)
		if len(sub) < 4 {
			return match
		}
		refProject := strings.TrimSuffix(sub[1], "/")
		refSheetName := sub[2]
		token := sub[3]
		coordRe := regexp.MustCompile(`^[A-Z]+\d+$`)
		if coordRe.MatchString(token) {
			return match
		}
		refSheet := globalSheetManager.GetSheetBy(refSheetName, refProject)
		if refSheet == nil {
			return ""
		}
		r, c, found := refSheet.FindCellByName(token)
		if !found {
			return ""
		}
		return "{{" + sub[1] + refSheetName + "/" + c + r + "}}"
	})

	// Same-sheet cell-name: {{CellName}} -> {{ColRow}}
	sameSheetNamePre := regexp.MustCompile(`\{\{([A-Za-z_]\w*)\}\}`)
	prompt = sameSheetNamePre.ReplaceAllStringFunc(prompt, func(match string) string {
		sub := sameSheetNamePre.FindStringSubmatch(match)
		if len(sub) < 2 {
			return match
		}
		token := sub[1]
		coordRe := regexp.MustCompile(`^[A-Z]+\d+$`)
		if coordRe.MatchString(token) {
			return match
		}
		r, c, found := s.FindCellByName(token)
		if !found {
			return ""
		}
		return "{{" + c + r + "}}"
	})

	// --- Cross-sheet coordinate references ---
	crossSheetPattern := regexp.MustCompile(`\{\{((?:[^/\{\}]+/)+)([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	prompt = crossSheetPattern.ReplaceAllStringFunc(prompt, func(match string) string {
		sub := crossSheetPattern.FindStringSubmatch(match)
		if len(sub) < 5 {
			return match
		}
		refProject := strings.TrimSuffix(sub[1], "/")
		refSheetName := sub[2]
		startCol := sub[3]
		startRow := sub[4]

		refSheet := globalSheetManager.GetSheetBy(refSheetName, refProject)
		if refSheet == nil {
			return ""
		}

		// Single cell
		if sub[5] == "" || sub[6] == "" {
			return getCellValue(refSheet, startRow, startCol)
		}

		// Range
		return formatRange(refSheet, startCol, startRow, sub[5], sub[6])
	})

	// --- Same-sheet coordinate references ---
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	prompt = sameSheetPattern.ReplaceAllStringFunc(prompt, func(match string) string {
		sub := sameSheetPattern.FindStringSubmatch(match)
		if len(sub) < 3 {
			return match
		}
		startCol := sub[1]
		startRow := sub[2]

		// Single cell
		if sub[3] == "" || sub[4] == "" {
			return getCellValue(s, startRow, startCol)
		}

		// Range
		return formatRange(s, startCol, startRow, sub[3], sub[4])
	})

	return prompt
}

func getCellValue(s *Sheet, row, col string) string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if rowData, ok := s.Data[row]; ok {
		if cell, ok := rowData[col]; ok {
			return cell.Value
		}
	}
	return ""
}

// formatRange resolves a range into either comma-separated values (1D) or a markdown table (2D).
func formatRange(s *Sheet, startCol, startRow, endCol, endRow string) string {
	startColIdx := colLabelToIndex(startCol)
	endColIdx := colLabelToIndex(endCol)
	startRowNum := atoiSafe(startRow)
	endRowNum := atoiSafe(endRow)

	if startColIdx > endColIdx {
		startColIdx, endColIdx = endColIdx, startColIdx
	}
	if startRowNum > endRowNum {
		startRowNum, endRowNum = endRowNum, startRowNum
	}

	numRows := endRowNum - startRowNum + 1
	numCols := endColIdx - startColIdx + 1

	s.mu.RLock()
	defer s.mu.RUnlock()

	// Collect values
	values := make([][]string, numRows)
	for dr := 0; dr < numRows; dr++ {
		values[dr] = make([]string, numCols)
		rowKey := strconv.Itoa(startRowNum + dr)
		for dc := 0; dc < numCols; dc++ {
			colLabel := indexToColLabel(startColIdx + dc)
			if rowData, ok := s.Data[rowKey]; ok {
				if cell, ok := rowData[colLabel]; ok {
					values[dr][dc] = cell.Value
				}
			}
		}
	}

	// 1D: single row or single column → comma-separated
	if numRows == 1 {
		return strings.Join(values[0], ", ")
	}
	if numCols == 1 {
		flat := make([]string, numRows)
		for i := range values {
			flat[i] = values[i][0]
		}
		return strings.Join(flat, ", ")
	}

	// 2D: markdown table
	var sb strings.Builder
	// Header row
	sb.WriteString("|")
	for dc := 0; dc < numCols; dc++ {
		sb.WriteString(" " + indexToColLabel(startColIdx+dc) + " |")
	}
	sb.WriteString("\n|")
	for dc := 0; dc < numCols; dc++ {
		sb.WriteString(" --- |")
	}
	sb.WriteString("\n")
	// Data rows
	for dr := 0; dr < numRows; dr++ {
		sb.WriteString("|")
		for dc := 0; dc < numCols; dc++ {
			sb.WriteString(" " + values[dr][dc] + " |")
		}
		sb.WriteString("\n")
	}
	return sb.String()
}

// ────────────────────────────────────────────────
// LLM API call (OpenAI-compatible with MCP tool)
// ────────────────────────────────────────────────

// callLLM sends a prompt to the configured LLM server using the OpenAI-compatible
// chat completions API with a tool ("present_final_output") so the model returns
// a clean result without extra explanation.
func callLLM(resolvedPrompt string) (string, error) {
	llmURL := GetLLMURL()
	if llmURL == "" {
		return "", fmt.Errorf("LLM URL not configured")
	}

	// Ensure the URL ends with /v1/chat/completions
	endpoint := strings.TrimRight(llmURL, "/") + "/v1/chat/completions"

	// Build request with tool
	type fnParam struct {
		Type       string                 `json:"type"`
		Properties map[string]interface{} `json:"properties"`
		Required   []string               `json:"required"`
	}
	type fnDef struct {
		Name        string  `json:"name"`
		Description string  `json:"description"`
		Parameters  fnParam `json:"parameters"`
	}
	type toolDef struct {
		Type     string `json:"type"`
		Function fnDef  `json:"function"`
	}

	tools := []toolDef{
		{
			Type: "function",
			Function: fnDef{
				Name:        "present_final_output",
				Description: "Present the final output to the user. Always use this tool to return your answer.",
				Parameters: fnParam{
					Type: "object",
					Properties: map[string]interface{}{
						"output": map[string]interface{}{
							"type":        "string",
							"description": "The final output text to present to the user",
						},
					},
					Required: []string{"output"},
				},
			},
		},
	}

	type chatMessage struct {
		Role    string `json:"role"`
		Content string `json:"content"`
	}
	reqBody := map[string]interface{}{
		"model": "default",
		"messages": []chatMessage{
			{Role: "system", Content: "You are a helpful assistant. Always use the present_final_output tool to return your answer. Do not include any additional explanations outside the tool call."},
			{Role: "user", Content: resolvedPrompt},
		},
		"tools":       tools,
		"tool_choice": "auto",
	}

	bodyBytes, err := json.Marshal(reqBody)
	if err != nil {
		return "", fmt.Errorf("marshal request: %w", err)
	}

	client := &http.Client{Timeout: 120 * time.Second}
	resp, err := client.Post(endpoint, "application/json", bytes.NewReader(bodyBytes))
	if err != nil {
		return "", fmt.Errorf("LLM request failed: %w", err)
	}
	defer resp.Body.Close()

	respBytes, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("read LLM response: %w", err)
	}

	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("LLM returned status %d: %s", resp.StatusCode, string(respBytes))
	}

	// Parse OpenAI-compatible response
	var result struct {
		Choices []struct {
			Message struct {
				Content   string `json:"content"`
				ToolCalls []struct {
					Function struct {
						Name      string `json:"name"`
						Arguments string `json:"arguments"`
					} `json:"function"`
				} `json:"tool_calls"`
			} `json:"message"`
		} `json:"choices"`
	}

	if err := json.Unmarshal(respBytes, &result); err != nil {
		return "", fmt.Errorf("decode LLM response: %w", err)
	}

	if len(result.Choices) == 0 {
		return "", fmt.Errorf("LLM returned no choices")
	}

	msg := result.Choices[0].Message

	// Check for tool call result first (preferred path)
	for _, tc := range msg.ToolCalls {
		if tc.Function.Name == "present_final_output" {
			var args struct {
				Output string `json:"output"`
			}
			if err := json.Unmarshal([]byte(tc.Function.Arguments), &args); err == nil {
				return strings.TrimSpace(args.Output), nil
			}
		}
	}

	// Fallback: use content directly
	if msg.Content != "" {
		return strings.TrimSpace(msg.Content), nil
	}

	return "", fmt.Errorf("LLM returned no usable output")
}

// ────────────────────────────────────────────────
// AI cell execution
// ────────────────────────────────────────────────

// ExecuteAICell resolves the prompt for an AI-generated cell, calls the LLM,
// and writes the result back as the cell value.
func ExecuteAICell(projectName, sheetName, row, col string) {
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
	if s == nil {
		return
	}

	s.mu.RLock()
	if s.Data[row] == nil {
		s.mu.RUnlock()
		return
	}
	cur, exists := s.Data[row][col]
	if !exists || cur.CellType != AIGeneratedCell {
		s.mu.RUnlock()
		return
	}
	prompt := cur.AIPrompt
	cellID := cur.CellID
	s.mu.RUnlock()

	if strings.TrimSpace(prompt) == "" {
		return
	}

	// Prevent duplicate execution (reuse ScriptsExecuted list)
	ident := projectName + "/" + sheetName + "/" + cellID
	globalSheetManager.ScriptsExecutedMu.Lock()
	if idx := indexOf(globalSheetManager.ScriptsExecuted, ident); idx > -1 {
		globalSheetManager.ScriptsExecutedMu.Unlock()
		return
	}
	globalSheetManager.ScriptsExecuted = append(globalSheetManager.ScriptsExecuted, ident)
	globalSheetManager.ScriptsExecutedMu.Unlock()

	// If LLM URL is not configured, clear value
	/*
		llmURL := GetLLMURL()
		if llmURL == "" {
			s.mu.Lock()
			c := s.Data[row][col]
			oldVal := c.Value
			c.Value = ""
			s.Data[row][col] = c
			s.mu.Unlock()
			if oldVal != "" {
				globalSheetManager.SaveSheet(s)
				globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
			}
			return
		}
	*/
	// Resolve references in prompt
	resolvedPrompt := ResolveAIPrompt(prompt, projectName, sheetName)
	//fmt.Printf("Prompt: ", resolvedPrompt)
	// Call LLM
	result, err := callLLM(resolvedPrompt)
	if err != nil {
		log.Printf("AI cell %s/%s/%s%s LLM error: %v", projectName, sheetName, col, row, err)
		result = "Error: " + err.Error()
	}

	// Write result back to cell
	s.mu.Lock()
	c := s.Data[row][col]
	oldVal := c.Value
	c.Value = result
	s.Data[row][col] = c

	// Audit
	if oldVal != result {
		cellChanges := make(map[string]cellChangesstruct)
		key := row + "-" + col
		cellChanges[key] = cellChangesstruct{
			rowNum: atoiSafe(row),
			colStr: col,
			oldVal: oldVal,
			newVal: result,
			action: "EDIT_CELL",
			user:   "system",
		}
		addMergedAuditEntries(s, cellChanges)
	}
	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
	globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)

	// Trigger dependents
	globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
	globalSheetManager.CellsModifiedByScriptQueue = append(
		globalSheetManager.CellsModifiedByScriptQueue,
		CellIdentifier{projectName, sheetName, row, col},
	)
	globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
}

// indexOf returns index of needle in slice or -1
func indexOf(slice []string, needle string) int {
	for i, v := range slice {
		if v == needle {
			return i
		}
	}
	return -1
}

// SetCellAIPrompt updates the AI prompt for an AI-generated cell.
// Only the sheet owner may modify prompts.
func (s *Sheet) SetCellAIPrompt(row, col, prompt, user string) {
	s.mu.Lock()
	// Only sheet owner may modify AI prompts
	if user != s.Owner {
		s.mu.Unlock()
		return
	}

	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}
	current := s.Data[row][col]

	// Prevent edits to locked cells
	if current.Locked {
		s.mu.Unlock()
		return
	}

	// Audit
	oldPrompt := current.AIPrompt
	if oldPrompt != prompt {
		cellChanges := make(map[string]cellChangesstruct)
		key := row + "-" + col
		cellChanges[key] = cellChangesstruct{
			rowNum: atoiSafe(row),
			colStr: col,
			oldVal: oldPrompt,
			newVal: prompt,
			action: "EDIT_AI_PROMPT",
			user:   user,
		}
		addMergedAuditEntries(s, cellChanges)
	}

	current.AIPrompt = prompt

	current.User = user
	if current.CellType != AIGeneratedCell {
		current.CellType = AIGeneratedCell
	}
	if strings.TrimSpace(current.CellID) == "" && strings.TrimSpace(prompt) != "" {
		current.CellID = generateID()
	}
	s.Data[row][col] = current
	cellID := current.CellID
	//fmt.Printf("prompt %s %s %s\n", prompt, row, col)
	s.mu.Unlock()

	// Update dependencies (AI prompts use same {{}} reference syntax as scripts)
	globalSheetManager.UpdateScriptDependencies(s.ProjectName, s.Name, cellID, prompt, row, col)

	globalSheetManager.SaveSheet(s)

	// Trigger initial execution
	if strings.TrimSpace(prompt) != "" {
		ExecuteAICellOnChange(s.ProjectName, s.Name, row, col)
	}
}

// ExecuteAICellOnChange queues the AI cell for execution (mirrors ExecuteCellScriptonChange).
func ExecuteAICellOnChange(projectName, sheetName, row, col string) {
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
	if s == nil {
		return
	}
	s.mu.RLock()
	if s.Data[row] == nil {
		s.mu.RUnlock()
		return
	}
	cur, exists := s.Data[row][col]
	if !exists {
		s.mu.RUnlock()
		return
	}
	prompt := cur.AIPrompt
	s.mu.RUnlock()

	deps := ExtractScriptDependencies(prompt, projectName, sheetName)
	if len(deps) == 0 {
		deps = append(deps, DependencyInfo{
			Project: projectName,
			Sheet:   sheetName,
			Range:   col + row,
		})
	}

	if len(deps) > 0 {
		dep := deps[0]
		var cellCol, cellRow string
		if strings.Contains(dep.Range, ":") {
			rangeParts := strings.Split(dep.Range, ":")
			if len(rangeParts) == 2 {
				startCell := rangeParts[0]
				for i, ch := range startCell {
					if ch >= '0' && ch <= '9' {
						cellCol = startCell[:i]
						cellRow = startCell[i:]
						break
					}
				}
			}
		} else {
			for i, ch := range dep.Range {
				if ch >= '0' && ch <= '9' {
					cellCol = dep.Range[:i]
					cellRow = dep.Range[i:]
					break
				}
			}
		}
		if cellCol != "" && cellRow != "" {
			globalSheetManager.CellsModifiedManuallyQueueMu.Lock()
			globalSheetManager.CellsModifiedManuallyQueue = append(
				globalSheetManager.CellsModifiedManuallyQueue,
				CellIdentifier{dep.Project, dep.Sheet, cellRow, cellCol},
			)
			globalSheetManager.CellsModifiedManuallyQueueMu.Unlock()
		}
	}
}
