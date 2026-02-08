package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"regexp"
	"slices"
	"strconv"
	"strings"
)

// Script Dependency Tracking System
//
// This system maintains a mapping between sheets and the scripts that depend on them.
// When a cell value changes, you can query which scripts need to be re-executed.

//
// Key Components:
// 1. ScriptIdentifier: Uniquely identifies a script and stores the referenced range
// 2. Dependency Map (in SheetManager): Maps sheet references to scripts that depend on them
//    - Example: "project1/sheet1" -> [{ScriptProjectName: "project1", ScriptSheetID: "sheet1", ScriptCellID: "B5cellid", ReferencedRange: "A2"}, ...]
//
// Usage Example:
//   // When a script is modified:
//   globalSheetManager.UpdateScriptDependencies(projectName, sheetID, cellID, script)
//
//   // When a cell value changes, get scripts that need re-execution:
//   dependents := globalSheetManager.GetDependentScripts(projectName, sheetID, row, col)
//   for _, dep := range dependents {
//       ExecuteCellScript(dep.ProjectName, dep.SheetID, depRow, depCol)
//   }
//
//   // Or use the helper function to execute all dependents:
//   ExecuteDependentScripts(projectName, sheetID, row, col)
//
// The system automatically updates dependencies when:
//   - Scripts are modified (via SetCellScript)
//   - Projects are renamed (via RenameProjectInDependencies)
//   - Sheets are loaded from disk (via rebuildScriptDependencies)
//   - Projects are renamed (via RenameProjectInDependencies)
//   - Sheets are loaded from disk (via rebuildScriptDependencies)

// ScriptIdentifier represents a unique script location
type ScriptIdentifier struct {
	ScriptProjectName string
	ScriptSheetID     string
	ScriptCellID      string
	ReferencedRange   string // e.g., "A2" or "A2:B3"
}

// String returns the string representation of the script identifier
func (si ScriptIdentifier) String() string {
	return si.ScriptProjectName + "/" + si.ScriptSheetID + "/" + si.ScriptCellID
}

// ParseScriptIdentifier parses a script identifier string
func ParseScriptIdentifier(s string) (ScriptIdentifier, bool) {
	parts := strings.Split(s, "/")
	if len(parts) != 3 {
		return ScriptIdentifier{}, false
	}
	return ScriptIdentifier{
		ScriptProjectName: parts[0],
		ScriptSheetID:     parts[1],
		ScriptCellID:      parts[2],
	}, true
}

// DependencyInfo represents a dependency with sheet and range information
type DependencyInfo struct {
	Project string
	Sheet   string
	Range   string
}

// ExtractScriptDependencies extracts all cell references from a script
// Returns a list of dependencies with project, sheet, and range information
// Examples:
//   - "{{A2}}" in script -> {Project: currentProject, Sheet: currentSheet, Range: "A2"}
//   - "{{A2:B3}}" in script -> {Project: currentProject, Sheet: currentSheet, Range: "A2:B3"}
//   - "{{project/sheet/A2}}" in script -> {Project: "project", Sheet: "sheet", Range: "A2"}
func ExtractScriptDependencies(script, currentProject, currentSheet string) []DependencyInfo {
	var deps []DependencyInfo
	seenDeps := make(map[string]bool)

	// Pattern to match {{projectname/sheetid/A2}} or {{projectname/sheetid/A2:B3}} (cross-sheet references)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+\d+(?::[A-Z]+\d+)?)\}\}`)

	// Pattern to match {{A2}} or {{A2:B3}} (same sheet references)
	tagPattern := regexp.MustCompile(`\{\{([A-Z]+\d+(?::[A-Z]+\d+)?)\}\}`)

	// Find all cross-sheet references
	matches := crossSheetPattern.FindAllStringSubmatch(script, -1)
	for _, match := range matches {
		if len(match) >= 4 {
			refProject := match[1]
			refSheet := match[2]
			refRange := match[3]
			key := refProject + "/" + refSheet + "/" + refRange
			if !seenDeps[key] {
				deps = append(deps, DependencyInfo{
					Project: refProject,
					Sheet:   refSheet,
					Range:   refRange,
				})
				seenDeps[key] = true
			}
		}
	}

	// Find all same-sheet references
	matches = tagPattern.FindAllStringSubmatch(script, -1)
	for _, match := range matches {
		if len(match) >= 2 {
			refRange := match[1]
			key := currentProject + "/" + currentSheet + "/" + refRange
			if !seenDeps[key] {
				deps = append(deps, DependencyInfo{
					Project: currentProject,
					Sheet:   currentSheet,
					Range:   refRange,
				})
				seenDeps[key] = true
			}
		}
	}

	return deps
}

// UpdateScriptDependencies updates the dependency map for a script
// Should be called whenever a script is modified
// projectName, sheetID, cellID identify the script location
// script is the new script content (used to extract dependencies)
// If script is empty, it will remove all dependencies for this script
// Function is used when script is modified
func (sm *SheetManager) UpdateScriptDependencies(scriptProjectName, scriptSheetID, scriptCellID, script string) {
	sm.scriptDepsMu.Lock()
	defer sm.scriptDepsMu.Unlock()

	// Remove old dependencies for this script (check all sheets)
	for sheetKey, scripts := range sm.scriptDeps {
		filtered := make([]ScriptIdentifier, 0, len(scripts))
		for _, s := range scripts {
			// Compare without ReferencedRange since we're identifying the script itself
			if s.ScriptProjectName != scriptProjectName || s.ScriptSheetID != scriptSheetID || s.ScriptCellID != scriptCellID {
				filtered = append(filtered, s)
			}
		}
		if len(filtered) > 0 {
			sm.scriptDeps[sheetKey] = filtered
		} else {
			delete(sm.scriptDeps, sheetKey)
		}
	}

	// If script is empty, we're done (dependencies removed)
	if strings.TrimSpace(script) == "" {
		return
	}

	// Add new dependencies
	deps := ExtractScriptDependencies(script, scriptProjectName, scriptSheetID)
	for _, dep := range deps {
		sheetKey := dep.Project + "/" + dep.Sheet
		scriptIdent := ScriptIdentifier{
			ScriptProjectName: scriptProjectName,
			ScriptSheetID:     scriptSheetID,
			ScriptCellID:      scriptCellID,
			ReferencedRange:   dep.Range,
		}
		sm.scriptDeps[sheetKey] = append(sm.scriptDeps[sheetKey], scriptIdent)
	}
}

// GetDependentScripts returns all scripts that depend on the given cell
// Checks if the cell matches any exact references or falls within range references
// If a script in the same cell depends on itself, it will be placed first in the result
func (sm *SheetManager) GetDependentScripts(projectName, sheetID, row, col string, skipSelf bool) []ScriptIdentifier {
	sm.scriptDepsMu.RLock()
	defer sm.scriptDepsMu.RUnlock()

	sheetKey := projectName + "/" + sheetID
	var result []ScriptIdentifier
	var sameCell *ScriptIdentifier // Track if there's a script in the same cell
	seen := make(map[string]bool)  // Track unique scripts: "project/sheet/cellid"

	// Get all scripts that depend on this sheet
	scripts, ok := sm.scriptDeps[sheetKey]
	if !ok {
		return result
	}

	// Construct cellID from row and col (e.g., "A2")
	cellID := col + row

	// Parse the changed cell
	cellRow := atoiSafe(row)
	cellColIdx := colLabelToIndex(col)

	// Check each script's referenced range
	for _, si := range scripts {
		scriptKey := si.ScriptProjectName + "/" + si.ScriptSheetID + "/" + si.ScriptCellID
		if seen[scriptKey] {
			continue
		}

		refRange := si.ReferencedRange
		matches := false

		// Check if it's a range reference (contains ":")
		if strings.Contains(refRange, ":") {
			// Parse range (e.g., "A2:B3")
			rangeParts := strings.Split(refRange, ":")
			if len(rangeParts) == 2 {
				startCell := rangeParts[0]
				endCell := rangeParts[1]

				// Parse start cell
				var startCol string
				var startRow int
				for i, ch := range startCell {
					if ch >= '0' && ch <= '9' {
						startCol = startCell[:i]
						startRow = atoiSafe(startCell[i:])
						break
					}
				}

				// Parse end cell
				var endCol string
				var endRow int
				for i, ch := range endCell {
					if ch >= '0' && ch <= '9' {
						endCol = endCell[:i]
						endRow = atoiSafe(endCell[i:])
						break
					}
				}

				startColIdx := colLabelToIndex(startCol)
				endColIdx := colLabelToIndex(endCol)

				// Check if cell is within range
				if cellRow >= startRow && cellRow <= endRow &&
					cellColIdx >= startColIdx && cellColIdx <= endColIdx {
					matches = true
				}
			}
		} else {
			// Single cell reference - exact match
			if refRange == cellID {
				matches = true
			}
		}

		if matches {
			// Check if this script is in the same cell
			if si.ScriptProjectName == projectName && si.ScriptSheetID == sheetID && si.ScriptCellID == cellID {
				// Store the same-cell script to add it first
				siCopy := si
				sameCell = &siCopy
			} else {
				result = append(result, si)
			}
			seen[scriptKey] = true
		}
	}

	// If there's a script in the same cell, prepend it to the result
	if sameCell != nil && !skipSelf {
		result = append([]ScriptIdentifier{*sameCell}, result...)
	}

	return result
}

// RenameProjectInDependencies updates all dependency references when a project is renamed
// It also updates the project name in all scripts that reference the old project name
func (sm *SheetManager) RenameProjectInDependencies(oldProject, newProject string) {
	sm.scriptDepsMu.Lock()
	defer sm.scriptDepsMu.Unlock()

	// Create new map with updated keys
	newDeps := make(map[string][]ScriptIdentifier)

	for sheetKey, scripts := range sm.scriptDeps {
		// Update sheet key if it starts with old project name
		newSheetKey := sheetKey
		parts := strings.Split(sheetKey, "/")
		if len(parts) >= 2 && parts[0] == oldProject {
			parts[0] = newProject
			newSheetKey = strings.Join(parts, "/")
		}

		// Update script identifiers
		newScripts := make([]ScriptIdentifier, 0, len(scripts))
		for _, script := range scripts {
			if script.ScriptProjectName == oldProject {
				script.ScriptProjectName = newProject
			}
			newScripts = append(newScripts, script)
		}

		newDeps[newSheetKey] = newScripts
	}

	sm.scriptDeps = newDeps

	// Update the sheets map keys and ProjectName field for sheets that belong to the renamed project
	sm.mu.Lock()
	sheetsToRemap := make([]*Sheet, 0)
	oldKeys := make([]string, 0)

	// First pass: identify sheets that need to be remapped
	for key, sheet := range sm.sheets {
		if sheet.ProjectName == oldProject {
			sheetsToRemap = append(sheetsToRemap, sheet)
			oldKeys = append(oldKeys, key)
		}
	}

	// Update sheet ProjectName and remap in sheets map
	for i, sheet := range sheetsToRemap {
		sheet.mu.Lock()
		sheet.ProjectName = newProject
		sheet.mu.Unlock()

		// Remove old key and add with new key
		delete(sm.sheets, oldKeys[i])
		sm.sheets[sheetKey(newProject, sheet.ID)] = sheet
	}
	sm.mu.Unlock()

	// Update project names in all scripts that contain cross-sheet references, regardless of
	// which project the script belongs to. Any script in any project may reference the renamed one.
	// Pattern: {{oldProject/sheetid/cellref}} -> {{newProject/sheetid/cellref}}
	sm.mu.RLock()
	sheetsToUpdate := make([]*Sheet, 0, len(sm.sheets))
	for _, sheet := range sm.sheets {
		sheetsToUpdate = append(sheetsToUpdate, sheet)
	}
	sm.mu.RUnlock()

	// Update scripts in all sheets
	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.Script) != "" {
					// Replace {{oldProject/...}} with {{newProject/...}} in scripts
					oldPattern := "{{" + oldProject + "/"
					newPattern := "{{" + newProject + "/"
					updatedScript := strings.ReplaceAll(cell.Script, oldPattern, newPattern)

					if updatedScript != cell.Script {
						// Must update the cell in the map (cell is a copy, not a reference)
						cellCopy := sheet.Data[rowKey][colKey]
						cellCopy.Script = updatedScript
						sheet.Data[rowKey][colKey] = cellCopy
						modified = true
					}
				}
			}
		}
		sheet.mu.Unlock()

		// Save the sheet if any scripts were modified
		if modified {
			sm.SaveSheet(sheet)

			// Send ROW_COL_UPDATED message to clients if sheet is opened
			if globalHub != nil {
				payload, _ := json.Marshal(sheet.SnapshotForClient())
				globalHub.broadcast <- &Message{
					Type:    "ROW_COL_UPDATED",
					SheetID: sheet.ID,
					Project: sheet.ProjectName,
					Payload: payload,
					User:    "system", // System update for project rename
				}
			}
		}
	}
}

// on script change
func ExecuteCellScriptonScriptChange(projectName, sheetID, row, col string) {
	//find a cell addressed in the script and add it in globalSheetManager.CellsModifiedManuallyQueue so that we can trigger script execution for that cell
	s := globalSheetManager.GetSheetBy(sheetID, projectName)
	if s == nil {
		return
	}
	if s.Data[row] == nil {
		return
	}
	cur, exists := s.Data[row][col]
	if !exists {
		return
	}

	// Extract all dependencies from the script
	deps := ExtractScriptDependencies(cur.Script, projectName, sheetID)

	// Add only the first referenced cell to the manually modified queue
	if len(deps) > 0 {
		dep := deps[0]

		// Parse the range to get the first cell
		var cellCol string
		var cellRow string

		if strings.Contains(dep.Range, ":") {
			// Range reference (e.g., "A2:B3") - take the first cell
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
			// Single cell reference (e.g., "A2")
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
				CellIdentifier{
					ProjectName: dep.Project,
					sheetID:     dep.Sheet,
					row:         cellRow,
					col:         cellCol,
				},
			)
			globalSheetManager.CellsModifiedManuallyQueueMu.Unlock()
		}
	}
}

// on row or column insert
func RefillingValuesonInsert(projectName, sheetID, row, col string) {

	WriteScriptOutputToCells(projectName, sheetID, row, col, false)

}

// ExecuteCellScript executes a Python script in a cell and updates the cell value
// with the script output. It handles tag replacement (e.g., {{A2}} or {{A2:B3}}),
// script execution, and populates cell spans if defined.
func ExecuteCellScript(projectName, sheetID, row, col string) {

	s := globalSheetManager.GetSheetBy(sheetID, projectName)
	if s == nil {
		return
	}

	// Read cell data with proper locking
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
	script := cur.Script
	rSpan := cur.ScriptOutput_RowSpan
	cSpan := cur.ScriptOutput_ColSpan
	cellID := cur.CellID

	if rSpan < 1 {
		rSpan = 1
	}
	if cSpan < 1 {
		cSpan = 1
	}

	// if globalSheetManager.scriptExecuted contains current script then return else add current script to globalSheetManager.scriptExecuted
	// Build the string identifier for the current script cell: "project/sheet/cellID"
	ident := projectName + "/" + sheetID + "/" + cellID

	// Check if already executed to prevent cycles
	globalSheetManager.ScriptsExecutedMu.Lock()
	indx := slices.Index(globalSheetManager.ScriptsExecuted, ident)

	if indx > -1 {
		//fmt.Println("Script already executed:", ident)
		globalSheetManager.ScriptsExecutedMu.Unlock()
		s.mu.RUnlock()
		return
	}
	s.mu.RUnlock()
	// Mark as executed for the duration of this call
	globalSheetManager.ScriptsExecuted = append(globalSheetManager.ScriptsExecuted, ident)
	globalSheetManager.ScriptsExecutedMu.Unlock()

	if strings.TrimSpace(script) == "" {
		s.mu.Lock()
		cur := s.Data[row][col]
		cur.ScriptOutput = ""
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		WriteScriptOutputToCells(projectName, sheetID, row, col, true)
		return
	}
	// Execute the script and update the cell value without logging EDIT_CELL
	ep := embeddedPy
	if ep == nil {
		// Store init error in the cell value
		cur := s.Data[row][col]
		if embeddedPyInitErr != nil {
			cur.Value = "Error: " + embeddedPyInitErr.Error()
		} else {
			cur.Value = "Error: Embedded Python not initialized"
		}
		s.mu.Lock()
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		return
	}
	// executing replaces the tags with values of the cells
	//example
	//1. if value at A2 is '7' then replace {{A2}} with '7' in the script
	//2. if tag is like {{A2:B3}} then replace with a 2D array string like [[7,8],[9,10]] from the cell values
	//3. For cell range from another sheet, e.g., {{projectname/sheetid/A2:B3}} , {{projectname/sheetid/B3}} etc

	// Pattern to match {{projectname/sheetid/A2}} or {{projectname/sheetid/A2:B3}} (cross-sheet references)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)

	// Pattern to match {{A2}} or {{A2:B3}} (same sheet references)
	tagPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)

	// First, handle cross-sheet references
	script = crossSheetPattern.ReplaceAllStringFunc(script, func(match string) string {
		submatches := crossSheetPattern.FindStringSubmatch(match)
		if len(submatches) < 5 {
			return match
		}

		refProjectName := submatches[1]
		refSheetID := submatches[2]
		startCol := submatches[3]
		startRow := submatches[4]

		// Get the referenced sheet
		refSheet := globalSheetManager.GetSheetBy(refSheetID, refProjectName)
		if refSheet == nil {
			return `""`
		}

		// Single cell reference {{projectname/sheetid/A2}}
		if submatches[5] == "" || submatches[6] == "" {
			refSheet.mu.RLock()
			defer refSheet.mu.RUnlock()

			if rowData, ok := refSheet.Data[startRow]; ok {
				if cell, ok := rowData[startCol]; ok {
					// Return the cell value - quote only if not a number
					val := cell.Value
					if _, err := strconv.ParseFloat(val, 64); err == nil {
						// It's a number, return unquoted
						return val
					}
					// Not a number, return quoted
					return fmt.Sprintf(`"%s"`, val)
				}
			}
			return `""`
		}

		// Range reference {{projectname/sheetid/A2:B3}}
		endCol := submatches[5]
		endRow := submatches[6]

		refSheet.mu.RLock()
		defer refSheet.mu.RUnlock()

		startColIdx := colLabelToIndex(startCol)
		endColIdx := colLabelToIndex(endCol)
		startRowNum := atoiSafe(startRow)
		endRowNum := atoiSafe(endRow)

		// Ensure proper order
		if startRowNum > endRowNum {
			startRowNum, endRowNum = endRowNum, startRowNum
		}
		if startColIdx > endColIdx {
			startColIdx, endColIdx = endColIdx, startColIdx
		}

		// Build 2D array
		var rows []string
		for r := startRowNum; r <= endRowNum; r++ {
			var cols []string
			for c := startColIdx; c <= endColIdx; c++ {
				rowKey := itoa(r)
				colLabel := indexToColLabel(c)

				val := ""
				if rowData, ok := refSheet.Data[rowKey]; ok {
					if cell, ok := rowData[colLabel]; ok {
						val = cell.Value
					}
				}
				// Quote only if not a number
				if _, err := strconv.ParseFloat(val, 64); err == nil && val != "" {
					// It's a number, use unquoted
					cols = append(cols, val)
				} else {
					// Not a number or empty, use quoted
					cols = append(cols, fmt.Sprintf(`"%s"`, val))
				}
			}
			rows = append(rows, "["+strings.Join(cols, ",")+"]")
		}

		return "[" + strings.Join(rows, ",") + "]"
	})

	// Then, handle same-sheet references
	script = tagPattern.ReplaceAllStringFunc(script, func(match string) string {
		submatches := tagPattern.FindStringSubmatch(match)
		if len(submatches) < 3 {
			return match
		}

		startCol := submatches[1]
		startRow := submatches[2]

		// Single cell reference {{A2}}
		if submatches[3] == "" || submatches[4] == "" {
			s.mu.RLock()
			defer s.mu.RUnlock()
			//cellKey := startRow + "-" + startCol
			if rowData, ok := s.Data[startRow]; ok {
				if cell, ok := rowData[startCol]; ok {
					// Return the cell value - quote only if not a number
					val := cell.Value
					if _, err := strconv.ParseFloat(val, 64); err == nil {
						// It's a number, return unquoted
						return val
					}
					// Not a number, return quoted
					return fmt.Sprintf(`"%s"`, val)
				}
			}
			return `""`
		}

		// Range reference {{A2:B3}}
		endCol := submatches[3]
		endRow := submatches[4]
		s.mu.RLock()
		defer s.mu.RUnlock()
		startColIdx := colLabelToIndex(startCol)
		endColIdx := colLabelToIndex(endCol)
		startRowNum := atoiSafe(startRow)
		endRowNum := atoiSafe(endRow)

		// Ensure proper order
		if startRowNum > endRowNum {
			startRowNum, endRowNum = endRowNum, startRowNum
		}
		if startColIdx > endColIdx {
			startColIdx, endColIdx = endColIdx, startColIdx
		}

		// Build 2D array
		var rows []string
		for r := startRowNum; r <= endRowNum; r++ {
			var cols []string
			for c := startColIdx; c <= endColIdx; c++ {
				rowKey := itoa(r)
				colLabel := indexToColLabel(c)

				val := ""
				if rowData, ok := s.Data[rowKey]; ok {
					if cell, ok := rowData[colLabel]; ok {
						val = cell.Value
					}
				}
				// Quote only if not a number
				if _, err := strconv.ParseFloat(val, 64); err == nil && val != "" {
					// It's a number, use unquoted
					cols = append(cols, val)
				} else {
					// Not a number or empty, use quoted
					cols = append(cols, fmt.Sprintf(`"%s"`, val))
				}
			}
			rows = append(rows, "["+strings.Join(cols, ",")+"]")
		}

		return "[" + strings.Join(rows, ",") + "]"
	})

	cmd, err := ep.PythonCmd("-c", script)
	if err != nil {
		s.mu.Lock()
		cur := s.Data[row][col]
		cur.Value = "Error: " + err.Error()
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		return
	}
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	runErr := cmd.Run()
	var newVal string
	if runErr != nil {
		errOut := strings.TrimRight(stderr.String(), "\r\n")
		if errOut == "" {
			errOut = runErr.Error()
		}
		newVal = "Error: " + errOut
	} else {
		newVal = strings.TrimRight(stdout.String(), "\r\n")
	}
	// Write ScriptOutput back and save
	s.mu.Lock()
	cur = s.Data[row][col] // Re-read to get latest state
	cur.ScriptOutput = newVal
	//fmt.Println("Script output for cell", cellID, ":", newVal)
	s.Data[row][col] = cur
	s.mu.Unlock()
	globalSheetManager.SaveSheet(s)

	// Call second function to write output to cell values
	WriteScriptOutputToCells(projectName, sheetID, row, col, true)
}

// WriteScriptOutputToCells writes the ScriptOutput to cell values
// Handles single values, arrays, and matrices based on the cell's span configuration
// triggernext will trigger script execution which depends on the cells being updated
// broadcastRowColUpdated sends a ROW_COL_UPDATED message for the given sheet
func broadcastRowColUpdated(s *Sheet, projectName, sheetID string) {
	if globalHub != nil {
		fmt.Println("Broadcasting ROW_COL_UPDATED for", projectName, sheetID, "after script execution")
		s.mu.RLock()
		payload, _ := json.Marshal(s.SnapshotForClient())
		s.mu.RUnlock()
		globalHub.broadcast <- &Message{
			Type:    "ROW_COL_UPDATED",
			SheetID: sheetID,
			Project: projectName,
			Payload: payload,
			User:    "system",
		}
	}
}

func WriteScriptOutputToCells(projectName, sheetID, row, col string, triggernext bool) {
	s := globalSheetManager.GetSheetBy(sheetID, projectName)
	if s == nil {
		return
	}

	s.mu.Lock()
	//defer s.mu.Unlock()

	if s.Data[row] == nil {
		return
	}
	cur, exists := s.Data[row][col]
	if !exists {
		return
	}

	newVal := cur.ScriptOutput
	rSpan := cur.ScriptOutput_RowSpan
	cSpan := cur.ScriptOutput_ColSpan

	// Prepare area: reset previous locked cells and validate target span area
	lockedbyScriptAt := "script-span " + cur.CellID
	// Resetting value and unlock previously locked cells belonging to this script-span
	cur.Value = ""
	for rKey, rowMap := range s.Data {
		for cKey, cell := range rowMap {
			if cell.Locked && cell.LockedBy == lockedbyScriptAt {
				cell.Value = ""
				cell.Locked = false
				cell.LockedBy = ""
				s.Data[rKey][cKey] = cell
			}
		}
	}

	if rSpan < 1 {
		rSpan = 1
	}
	if cSpan < 1 {
		cSpan = 1
	}

	// Validate emptiness for spanned area (excluding top-left)
	baseIdx := colLabelToIndex(col)
	baseRow := atoiSafe(row)
	if rSpan > 1 || cSpan > 1 {
		blocked := false
		for dr := 0; dr < rSpan && !blocked; dr++ {
			rKey := itoa(baseRow + dr)
			for dc := 0; dc < cSpan; dc++ {
				if dr == 0 && dc == 0 {
					continue
				}
				cLabel := indexToColLabel(baseIdx + dc)
				cell, ok := s.Data[rKey][cLabel]
				if ok && (strings.TrimSpace(cell.Value) != "" || strings.TrimSpace(cell.Script) != "") {
					blocked = true
					break
				}
			}
		}
		if blocked {
			rSpan = 1
			cSpan = 1
		}
		// Apply spans on top-left cell now that area is clear
		cur.ScriptOutput_RowSpan = rSpan
		cur.ScriptOutput_ColSpan = cSpan
		// Lock covered cells
		for dr := 0; dr < rSpan; dr++ {
			rKey := itoa(baseRow + dr)
			if s.Data[rKey] == nil {
				s.Data[rKey] = make(map[string]Cell)
			}
			for dc := 0; dc < cSpan; dc++ {
				if dr == 0 && dc == 0 {
					continue
				}
				cLabel := indexToColLabel(baseIdx + dc)
				c := s.Data[rKey][cLabel]
				c.Locked = true
				c.LockedBy = lockedbyScriptAt
				s.Data[rKey][cLabel] = c
			}
		}
	} else {
		// Ensure spans are set even for 1x1
		cur.ScriptOutput_RowSpan = 1
		cur.ScriptOutput_ColSpan = 1
	}
	// Persist any updates to the top-left cell
	s.Data[row][col] = cur

	// If no span, simply set value
	if rSpan == 1 && cSpan == 1 {
		cur.Value = newVal
		s.Data[row][col] = cur
		if triggernext {
			globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
			globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, row, col})
			globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
		}
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		broadcastRowColUpdated(s, projectName, sheetID)
		return
	}

	// baseIdx and baseRow already computed above

	// Try to parse as single value (string, number, boolean)
	// If it's not a JSON array or matrix, just set the value to the base cell
	var testInterface interface{}
	if err := json.Unmarshal([]byte(newVal), &testInterface); err != nil {
		// Not valid JSON, treat as plain string
		cur.Value = newVal
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		broadcastRowColUpdated(s, projectName, sheetID)
		return
	}

	// Check if it's a simple value (not array or object)
	switch testInterface.(type) {
	case string, float64, bool, nil:
		// Simple value, set to base cell
		cur.Value = newVal
		s.Data[row][col] = cur
		if triggernext {
			globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
			globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, row, col})
			globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
		}
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		broadcastRowColUpdated(s, projectName, sheetID)
		return
	}

	// handling matrix type [[1,2],[3,4]]
	var matrix [][]string
	MatrixParsed := false
	if rSpan > 1 || cSpan > 1 {
		var any interface{}
		if err := json.Unmarshal([]byte(newVal), &any); err == nil {
			if arr, ok := any.([]interface{}); ok {
				tmp := make([][]string, 0, len(arr))
				for _, r := range arr {
					if rowArr, ok := r.([]interface{}); ok {
						rowVals := make([]string, 0, len(rowArr))
						for _, v := range rowArr {
							rowVals = append(rowVals, fmt.Sprint(v))
						}
						tmp = append(tmp, rowVals)
					}
				}
				if len(tmp) > 0 {
					matrix = tmp
					MatrixParsed = true
				}
			}
		}
	}

	if MatrixParsed {
		// if matrix dimensions is more than ScriptOutput_RowSpan x ScriptOutput_ColSpan, we will simply fill cur.Value = scriptOutput
		if len(matrix) > rSpan || (len(matrix) > 0 && len(matrix[0]) > cSpan) {
			cur.Value = cur.ScriptOutput
			s.Data[row][col] = cur
			s.mu.Unlock()
			globalSheetManager.SaveSheet(s)
			broadcastRowColUpdated(s, projectName, sheetID)
			return
		}
		//else

		for dr := 0; dr < rSpan; dr++ {
			rKey := itoa(baseRow + dr)
			if s.Data[rKey] == nil {
				s.Data[rKey] = make(map[string]Cell)
			}
			for dc := 0; dc < cSpan; dc++ {
				cLabel := indexToColLabel(baseIdx + dc)
				var val string
				if dr < len(matrix) && dc < len(matrix[dr]) {
					val = matrix[dr][dc]
				}
				c := s.Data[rKey][cLabel]
				c.Value = val
				s.Data[rKey][cLabel] = c
				if triggernext {

					globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
					globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, rKey, cLabel})
					globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()

				}
			}
		}

		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		broadcastRowColUpdated(s, projectName, sheetID)
		return
	}

	// handling array type [1,2,3,4]
	// If output is a flat array and span is 1xN or Nx1, fill accordingly
	var arrAny []interface{}
	if err := json.Unmarshal([]byte(cur.ScriptOutput), &arrAny); err == nil {
		if (rSpan == 1 && cSpan > 1 && len(arrAny) <= cSpan) || (cSpan == 1 && rSpan > 1 && len(arrAny) <= rSpan) {
			for i, v := range arrAny {
				if rSpan == 1 {
					// Fill row left to right
					cLabel := indexToColLabel(baseIdx + i)
					c := s.Data[row][cLabel]
					c.Value = fmt.Sprint(v)
					s.Data[row][cLabel] = c
					if triggernext {
						globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
						globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, row, cLabel})
						globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
					}

				} else {
					// Fill column top to bottom
					rKey := itoa(baseRow + i)
					c := s.Data[rKey][col]
					c.Value = fmt.Sprint(v)
					s.Data[rKey][col] = c
					if triggernext {
						globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
						globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, rKey, col})
						globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
					}
				}
			}
		} else {
			cur.Value = cur.ScriptOutput
			s.Data[row][col] = cur
			if triggernext {
				globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
				globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetID, row, col})
				globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
			}
		}
	}
	s.mu.Unlock()
	globalSheetManager.SaveSheet(s)

	broadcastRowColUpdated(s, projectName, sheetID)
}

// ExecuteDependentScripts executes all scripts that depend on the given cell
// This should be called when a cell value is modified to trigger cascading updates
// skipSelf prevent execution of the script in the same cell
func ExecuteDependentScripts(projectName, sheetID, row, col string, skipSelf bool) {
	// Get the cell ID for the modified cell
	s := globalSheetManager.GetSheetBy(sheetID, projectName)
	if s == nil {
		return
	}

	// Get all dependent scripts
	dependents := globalSheetManager.GetDependentScripts(projectName, sheetID, row, col, skipSelf)

	// Execute each dependent script
	for _, dep := range dependents {
		//fmt.Println("Executing dependent script at ", dep.ScriptProjectName, "/", dep.ScriptSheetID, " cell ", dep.ScriptCellID, " which depends on ", projectName, "/", sheetID, " cell ", row+col)
		ExecuteCellScriptWithIdentifier(dep)
	}
}

func ExecuteCellScriptWithIdentifier(dep ScriptIdentifier) {

	// Find the cell with this script
	depSheet := globalSheetManager.GetSheetBy(dep.ScriptSheetID, dep.ScriptProjectName)
	if depSheet == nil {
		return
	}

	// Find the cell by CellID
	depSheet.mu.RLock()
	var depRow, depCol string
	found := false
	for r, rowMap := range depSheet.Data {
		for c, cell := range rowMap {
			if cell.CellID == dep.ScriptCellID {
				depRow = r
				depCol = c
				found = true
				break
			}
		}
		if found {
			break
		}
	}
	depSheet.mu.RUnlock()

	if found {
		// Execute the dependent script
		ExecuteCellScript(dep.ScriptProjectName, dep.ScriptSheetID, depRow, depCol)
	}

}
