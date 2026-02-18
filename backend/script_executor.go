package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"regexp"
	"slices"
	"strconv"
	"strings"
	"time"
)

// ScriptIdentifier represents a unique script location
type ScriptIdentifier struct {
	ScriptProjectName string
	ScriptSheetName   string
	ScriptCellID      string
	ReferencedRange   string // e.g., "A2" or "A2:B3" is the cell or range that this script depends on (used for dependency tracking)
	// projectname and sheetName in ReferencedRange are not needed because they are already captured in the key used in scriptDeps map[string][]ScriptIdentifier
}

type cellChangesstruct struct {
	rowNum int
	colStr string
	oldVal string
	newVal string
	action string
	user   string
}

/*
cellChanges := make(map[string]struct {
		rowNum int
		colStr string
		oldVal string
		newVal string
		action string
		user   string
	})
*/
// String returns the string representation of the script identifier
func (si ScriptIdentifier) String() string {
	return si.ScriptProjectName + "/" + si.ScriptSheetName + "/" + si.ScriptCellID
}

// ParseScriptIdentifier parses a script identifier string
func ParseScriptIdentifier(s string) (ScriptIdentifier, bool) {
	parts := strings.Split(s, "/")
	if len(parts) != 3 {
		return ScriptIdentifier{}, false
	}
	return ScriptIdentifier{
		ScriptProjectName: parts[0],
		ScriptSheetName:   parts[1],
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

	// Pattern to match {{project/.../sheetid/A2}} or {{project/.../sheetid/A2:B3}} (cross-sheet references).
	// The project path may contain slashes (subfolder), so we allow any characters except {{ and }} before
	// the last two slash-delimited segments (sheetid and cell range).
	crossSheetPattern := regexp.MustCompile(`\{\{((?:[^/\{\}]+/)+)([^/\{\}]+)/([A-Z]+\d+(?::[A-Z]+\d+)?)\}\}`)

	// Pattern to match {{A2}} or {{A2:B3}} (same sheet references)
	tagPattern := regexp.MustCompile(`\{\{([A-Z]+\d+(?::[A-Z]+\d+)?)\}\}`)

	// Find all cross-sheet references
	matches := crossSheetPattern.FindAllStringSubmatch(script, -1)
	for _, match := range matches {
		if len(match) >= 4 {
			// match[1] is "project/.../" (trailing slash), match[2] is sheetid, match[3] is range
			refProject := strings.TrimSuffix(match[1], "/")
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
// projectName, sheetName, cellID identify the script location
// script is the new script content (used to extract dependencies)
// If script is empty, it will remove all dependencies for this script
// Function is used when script is modified
// eg:- scriptDeps["projectName/sheetName"] = [ScriptIdentifier{ScriptProjectName: projectName, ScriptSheetName: sheetName, ScriptCellID: cellID, ReferencedRange: "A2"})
func (sm *SheetManager) UpdateScriptDependencies(scriptProjectName, scriptSheetName, scriptCellID, script, rowLabel, colLabel string) {
	sm.scriptDepsMu.Lock()
	defer sm.scriptDepsMu.Unlock()

	// Remove old dependencies for this script (check all sheets)
	for sheetKey, scripts := range sm.scriptDeps {
		filtered := make([]ScriptIdentifier, 0, len(scripts))
		for _, s := range scripts {
			// Compare without ReferencedRange since we're identifying the script itself
			if s.ScriptProjectName != scriptProjectName || s.ScriptSheetName != scriptSheetName || s.ScriptCellID != scriptCellID {
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
		s := globalSheetManager.GetSheetBy(scriptSheetName, scriptProjectName)
		lockedbyScriptAt := "script-span " + scriptCellID
		for rKey, rowMap := range s.Data {
			for cKey, cell := range rowMap {
				if cell.Locked && cell.LockedBy == lockedbyScriptAt {
					cell.Value = ""
					cell.Value_FromNonSelfScript = ""
					cell.Locked = false
					cell.LockedBy = ""
					s.Data[rKey][cKey] = cell
				}
			}
		}
		globalSheetManager.SaveSheet(s)
		return
	}

	// Add new dependencies
	deps := ExtractScriptDependencies(script, scriptProjectName, scriptSheetName)
	if len(deps) == 0 {
		// No explicit references found; add self cell as dependency
		//fmt.Println("No explicit references found in script, adding self reference for cell ", scriptCellID)
		if strings.TrimSpace(colLabel) != "" && strings.TrimSpace(rowLabel) != "" {
			deps = append(deps, DependencyInfo{
				Project: scriptProjectName,
				Sheet:   scriptSheetName,
				Range:   colLabel + rowLabel,
			})
		}
	}
	for _, dep := range deps {
		sheetKey := dep.Project + "/" + dep.Sheet
		scriptIdent := ScriptIdentifier{
			ScriptProjectName: scriptProjectName,
			ScriptSheetName:   scriptSheetName,
			ScriptCellID:      scriptCellID,
			ReferencedRange:   dep.Range,
		}
		sm.scriptDeps[sheetKey] = append(sm.scriptDeps[sheetKey], scriptIdent)
	}
}

// GetDependentScripts returns all scripts that depend on the given cell
// Checks if the cell matches any exact references or falls within range references
// If a script in the same cell depends on itself, it will be placed first in the result
func (sm *SheetManager) GetDependentScripts(projectName, sheetName, row, col string) []ScriptIdentifier {
	sm.scriptDepsMu.RLock()
	defer sm.scriptDepsMu.RUnlock()

	sheetKey := projectName + "/" + sheetName
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
		scriptKey := si.ScriptProjectName + "/" + si.ScriptSheetName + "/" + si.ScriptCellID
		if seen[scriptKey] {
			continue
		}
		//fmt.Println("Checking script at ", si.ScriptProjectName, "/", si.ScriptSheetName, " cell ", si.ScriptCellID, " with reference ", si.ReferencedRange, " against changed cell ", projectName, "/", sheetName, " cell ", col+row)
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
			if si.ScriptProjectName == projectName && si.ScriptSheetName == sheetName && si.ScriptCellID == cellID {
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
	if sameCell != nil {
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

	// First pass: identify sheets that need to be remapped (including sheets in subfolders)
	for key, sheet := range sm.sheets {
		if sheet.ProjectName == oldProject || strings.HasPrefix(sheet.ProjectName, oldProject+"/") {
			sheetsToRemap = append(sheetsToRemap, sheet)
			oldKeys = append(oldKeys, key)
		}
	}

	// Update sheet ProjectName and remap in sheets map
	for i, sheet := range sheetsToRemap {
		sheet.mu.Lock()
		var updatedPN string
		if sheet.ProjectName == oldProject {
			updatedPN = newProject
		} else {
			// subfolder: replace only the leading oldProject prefix
			updatedPN = newProject + sheet.ProjectName[len(oldProject):]
		}
		sheet.ProjectName = updatedPN
		sheet.mu.Unlock()

		// Remove old key and add with new key
		delete(sm.sheets, oldKeys[i])
		sm.sheets[sheetKey(updatedPN, sheet.Name)] = sheet
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
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			// Send ROW_COL_UPDATED message to clients if sheet is opened
			/*
				if globalHub != nil {
					payload, _ := json.Marshal(sheet.SnapshotForClient())
					globalHub.broadcast <- &Message{
						Type:    "ROW_COL_UPDATED",
						SheetName: sheet.ID,
						Project: sheet.ProjectName,
						Payload: payload,
						User:    "system", // System update for project rename
					}
				}
			*/
		}
	}
}

// RenameSheetInDependencies updates all dependency references when a sheet is renamed within a project.
// It also updates the sheet name in all scripts that reference the old sheet name.
func (sm *SheetManager) RenameSheetInDependencies(projectName, oldSheetName, newSheetName string) {
	sm.scriptDepsMu.Lock()
	defer sm.scriptDepsMu.Unlock()

	// Create new map with updated keys
	newDeps := make(map[string][]ScriptIdentifier)

	for sk, scripts := range sm.scriptDeps {
		// Update sheet key if it matches the renamed sheet
		newSk := sk
		parts := strings.Split(sk, "/")
		if len(parts) >= 2 && parts[0] == projectName && parts[1] == oldSheetName {
			parts[1] = newSheetName
			newSk = strings.Join(parts, "/")
		}

		// Update script identifiers
		newScripts := make([]ScriptIdentifier, 0, len(scripts))
		for _, script := range scripts {
			if script.ScriptProjectName == projectName && script.ScriptSheetName == oldSheetName {
				script.ScriptSheetName = newSheetName
			}
			newScripts = append(newScripts, script)
		}

		newDeps[newSk] = newScripts
	}

	sm.scriptDeps = newDeps

	// Update sheet names in all scripts that contain cross-sheet references to the renamed sheet.
	// Pattern: {{projectName/oldSheetName/cellref}} -> {{projectName/newSheetName/cellref}}
	sm.mu.RLock()
	sheetsToUpdate := make([]*Sheet, 0, len(sm.sheets))
	for _, sheet := range sm.sheets {
		sheetsToUpdate = append(sheetsToUpdate, sheet)
	}
	sm.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.Script) != "" {
					oldPattern := "{{" + projectName + "/" + oldSheetName + "/"
					newPattern := "{{" + projectName + "/" + newSheetName + "/"
					updatedScript := strings.ReplaceAll(cell.Script, oldPattern, newPattern)

					if updatedScript != cell.Script {
						cellCopy := sheet.Data[rowKey][colKey]
						cellCopy.Script = updatedScript
						sheet.Data[rowKey][colKey] = cellCopy
						modified = true
					}
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sm.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

// on script change
func ExecuteCellScriptonChange(projectName, sheetName, row, col string) {
	//find a cell addressed in the script and add it in globalSheetManager.CellsModifiedManuallyQueue so that we can trigger script execution for that cell
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
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
	deps := ExtractScriptDependencies(cur.Script, projectName, sheetName)
	if len(deps) == 0 {
		// No explicit references found; add self cell as dependency
		deps = append(deps, DependencyInfo{
			Project: projectName,
			Sheet:   sheetName,
			Range:   col + row,
		})
	}
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
			//Add in the queue to execute other wise mutex will lock up
			globalSheetManager.CellsModifiedManuallyQueueMu.Lock()
			globalSheetManager.CellsModifiedManuallyQueue = append(
				globalSheetManager.CellsModifiedManuallyQueue,
				CellIdentifier{
					ProjectName: dep.Project,
					sheetName:   dep.Sheet,
					row:         cellRow,
					col:         cellCol,
				},
			)
			globalSheetManager.CellsModifiedManuallyQueueMu.Unlock()
		}
	}
}

// CheckIfScriptReferencesSelf checks if a script references its own cell
// Returns true if the script depends on the cell where it's located
func CheckIfScriptReferencesSelf(script, projectName, sheetName, cellID string) bool {
	deps := ExtractScriptDependencies(script, projectName, sheetName)
	for _, dep := range deps {
		if dep.Project == projectName && dep.Sheet == sheetName {
			// Check if the reference matches the cell ID
			if strings.Contains(dep.Range, ":") {
				// Range reference - check if cellID falls within range
				rangeParts := strings.Split(dep.Range, ":")
				if len(rangeParts) == 2 {
					startCell := rangeParts[0]
					endCell := rangeParts[1]

					// Parse cellID
					var cellCol string
					var cellRow int
					for i, ch := range cellID {
						if ch >= '0' && ch <= '9' {
							cellCol = cellID[:i]
							cellRow = atoiSafe(cellID[i:])
							break
						}
					}

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

					cellColIdx := colLabelToIndex(cellCol)
					startColIdx := colLabelToIndex(startCol)
					endColIdx := colLabelToIndex(endCol)

					// Check if cell is within range
					if cellRow >= startRow && cellRow <= endRow &&
						cellColIdx >= startColIdx && cellColIdx <= endColIdx {
						return true
					}
				}
			} else {
				// Single cell reference
				if dep.Range == cellID {
					return true
				}
			}
		}
	}
	return false
}

// ExecuteCellScript executes a Python script in a cell and updates the cell value
// with the script output. It handles tag replacement (e.g., {{A2}} or {{A2:B3}}),
// script execution, and populates cell spans if defined.
func ExecuteCellScript(projectName, sheetName, row, col string) {

	s := globalSheetManager.GetSheetBy(sheetName, projectName)
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
	ident := projectName + "/" + sheetName + "/" + cellID
	//fmt.Println("script for cell", cellID, "with content:", script)
	// Check if already executed to prevent cycles
	// Release sheet read lock before acquiring global scripts-executed lock to avoid nested locks
	s.mu.RUnlock()
	globalSheetManager.ScriptsExecutedMu.Lock()
	indx := slices.Index(globalSheetManager.ScriptsExecuted, ident)

	if indx > -1 {
		//fmt.Println("Script already executed:", ident)
		globalSheetManager.ScriptsExecutedMu.Unlock()
		return
	}
	// Mark as executed for the duration of this call
	globalSheetManager.ScriptsExecuted = append(globalSheetManager.ScriptsExecuted, ident)
	globalSheetManager.ScriptsExecutedMu.Unlock()
	//fmt.Println("executing cell script:", ident)
	if strings.TrimSpace(script) == "" {
		s.mu.Lock()
		cur := s.Data[row][col]
		cur.ScriptOutput = ""
		cur.ScriptOutput_RowSpan = 1
		cur.ScriptOutput_ColSpan = 1
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		WriteScriptOutputToCells(projectName, sheetName, row, col, true, false)
		return
	}
	// Execute the script and update the cell value
	ep := embeddedPy
	if ep == nil {
		// Store init error in the cell value
		cur := s.Data[row][col]
		if embeddedPyInitErr != nil {
			cur.Value = "Error: " + embeddedPyInitErr.Error()
		} else {
			cur.Value = "Error: Embedded Python not initialized"
		}
		cur.ScriptOutput_RowSpan = 1
		cur.ScriptOutput_ColSpan = 1
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

	// Pattern to match {{project/.../sheetid/A2}} or {{project/.../sheetid/A2:B3}} (cross-sheet references).
	// Project path may contain slashes (subfolder support).
	crossSheetPattern := regexp.MustCompile(`\{\{((?:[^/\{\}]+/)+)([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)

	// Pattern to match {{A2}} or {{A2:B3}} (same sheet references)
	tagPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)

	// First, handle cross-sheet references
	script = crossSheetPattern.ReplaceAllStringFunc(script, func(match string) string {
		submatches := crossSheetPattern.FindStringSubmatch(match)
		if len(submatches) < 5 {
			return match
		}

		// submatches[1] is "project/.../" (trailing slash), submatches[2] is sheetid
		refProjectName := strings.TrimSuffix(submatches[1], "/")
		refSheetName := submatches[2]
		startCol := submatches[3]
		startRow := submatches[4]

		// Get the referenced sheet
		refSheet := globalSheetManager.GetSheetBy(refSheetName, refProjectName)
		if refSheet == nil {
			return `""`
		}

		// Single cell reference {{projectname/sheetid/A2}}
		if submatches[5] == "" || submatches[6] == "" {
			refSheet.mu.RLock()
			defer refSheet.mu.RUnlock()

			if rowData, ok := refSheet.Data[startRow]; ok {
				if cell, ok := rowData[startCol]; ok {
					// Return the cell value - numbers unquoted; strings properly escaped
					val := cell.Value
					if val == "" {
						return `""`
					}
					if _, err := strconv.ParseFloat(val, 64); err == nil {
						// It's a number, return unquoted
						return val
					}
					// Not a number, return quoted with escapes safe for Python
					return strconv.Quote(val)
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
				// Numbers unquoted; strings properly escaped, empty as ""
				if _, err := strconv.ParseFloat(val, 64); err == nil && val != "" {
					cols = append(cols, val)
				} else {
					cols = append(cols, strconv.Quote(val))
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
			if rowData, ok := s.Data[startRow]; ok {
				if cell, ok := rowData[startCol]; ok {
					// Check if this reference is to the same cell as the script location
					refCellID := startCol + startRow
					var val string
					if refCellID == cellID {
						// Self-reference: use Value_FromNonSelfScript
						val = cell.Value_FromNonSelfScript
					} else {
						// Different cell: use Value
						val = cell.Value
					}
					if val == "" {
						return `""`
					}
					if _, err := strconv.ParseFloat(val, 64); err == nil {
						// It's a number, return unquoted
						return val
					}
					// Not a number, return quoted with escapes safe for Python
					return strconv.Quote(val)
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
						// Check if this cell in the range is the same as the script location
						refCellID := colLabel + rowKey
						if refCellID == cellID {
							// Self-reference: use Value_FromNonSelfScript
							val = cell.Value_FromNonSelfScript
						} else {
							// Different cell: use Value
							val = cell.Value
						}
					}
				}
				// Numbers unquoted; strings properly escaped, empty as ""
				if _, err := strconv.ParseFloat(val, 64); err == nil && val != "" {
					cols = append(cols, val)
				} else {
					cols = append(cols, strconv.Quote(val))
				}
			}
			rows = append(rows, "["+strings.Join(cols, ",")+"]")
		}

		return "[" + strings.Join(rows, ",") + "]"
	})

	cmd, err := ep.PythonCmd("-c", script)
	//fmt.Println("Executing script ", script)
	if err != nil {
		s.mu.Lock()
		cur := s.Data[row][col]
		cur.Value = "Error: " + err.Error()
		cur.ScriptOutput_RowSpan = 1
		cur.ScriptOutput_ColSpan = 1
		s.Data[row][col] = cur
		s.mu.Unlock()
		globalSheetManager.SaveSheet(s)
		WriteScriptOutputToCells(projectName, sheetName, row, col, true, true)
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

	// Check if script references its own cell
	isSelfReferencing := CheckIfScriptReferencesSelf(script, projectName, sheetName, cellID)

	// Call second function to write output to cell values
	WriteScriptOutputToCells(projectName, sheetName, row, col, true, isSelfReferencing)
}

// addMergedAuditEntries adds audit entries for cell changes, merging with existing entries from the same user within 24 hours
func addMergedAuditEntries(s *Sheet, cellChanges map[string]cellChangesstruct) {
	now := time.Now()
	for _, change := range cellChanges {
		// Find the latest matching edit for this cell by user "system"
		prevIdx := -1
		for i := len(s.AuditLog) - 1; i >= 0; i-- {
			entry := s.AuditLog[i]
			if entry.Action == change.action && entry.Row1 == change.rowNum && entry.Col1 == change.colStr {
				if entry.User == change.user && !entry.ChangeReversed {
					prevIdx = i
				}
				break
			}
		}

		oldValForNew := change.oldVal
		if prevIdx >= 0 {
			// Only merge if previous log is within 24 hours
			if time.Since(s.AuditLog[prevIdx].Timestamp) < 24*time.Hour {
				oldValForNew = s.AuditLog[prevIdx].OldValue
				s.AuditLog = append(s.AuditLog[:prevIdx], s.AuditLog[prevIdx+1:]...)
			}
		}

		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp:      now,
			User:           change.user,
			Action:         change.action,
			Row1:           change.rowNum,
			Col1:           change.colStr,
			OldValue:       oldValForNew,
			NewValue:       change.newVal,
			ChangeReversed: false,
		})
	}
}

// This function causes mutex lockout when called concurrently (observed in tests) - needs refactor to avoid nested locks and long lock durations
func WriteScriptOutputToCells(projectName, sheetName, row, col string, triggernext bool, isSelfReferencing bool) {
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
	if s == nil {
		return
	}

	s.mu.Lock()
	//defer s.mu.Unlock()

	if s.Data[row] == nil {
		s.mu.Unlock()
		return
	}
	cur, exists := s.Data[row][col]
	if !exists {
		s.mu.Unlock()
		return
	}

	newVal := cur.ScriptOutput

	// Helper: normalize Python-style lists (single quotes, True/False, None) to JSON
	normalizePythonListToJSON := func(s string) string {
		if strings.Contains(s, "[") && strings.Contains(s, "]") {
			reTrue := regexp.MustCompile(`\bTrue\b`)
			reFalse := regexp.MustCompile(`\bFalse\b`)
			reNone := regexp.MustCompile(`\bNone\b`)
			s = reTrue.ReplaceAllString(s, "true")
			s = reFalse.ReplaceAllString(s, "false")
			s = reNone.ReplaceAllString(s, "null")
			// Replace single-quoted strings with double-quoted strings
			reSingle := regexp.MustCompile(`'([^'\\]*)'`)
			s = reSingle.ReplaceAllString(s, `"$1"`)
		}
		return s
	}
	rSpan := cur.ScriptOutput_RowSpan
	cSpan := cur.ScriptOutput_ColSpan

	// Prepare area: reset previous locked cells and validate target span area
	lockedbyScriptAt := "script-span " + cur.CellID

	// Capture all previous values before clearing (map key: "row-col")
	previousValues := make(map[string]string)
	previousValues[row+"-"+col] = cur.Value

	// Resetting value and unlock previously locked cells belonging to this script-span
	cur.Value = ""
	if !isSelfReferencing {
		cur.Value_FromNonSelfScript = ""
	}
	for rKey, rowMap := range s.Data {
		for cKey, cell := range rowMap {
			if cell.Locked && cell.LockedBy == lockedbyScriptAt {
				previousValues[rKey+"-"+cKey] = cell.Value
				cell.Value = ""
				cell.Value_FromNonSelfScript = ""
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

	// Track all cell changes for audit logging - map key: "row-col"
	cellChanges := make(map[string]cellChangesstruct)
	// Helper to record a cell change
	recordChange := func(rKey, cLabel, oldValue, newValue string) {
		if oldValue != newValue {
			key := rKey + "-" + cLabel
			cellChanges[key] = cellChangesstruct{
				rowNum: atoiSafe(rKey),
				colStr: cLabel,
				oldVal: oldValue,
				newVal: newValue,
				action: "EDIT_CELL",
				user:   "system",
			}
		}
	}

	// If no span, simply set value
	if rSpan == 1 && cSpan == 1 {
		oldVal := previousValues[row+"-"+col]
		if isSelfReferencing {
			// Script references itself - use Value_FromNonSelfScript in the script
			cur.Value = cur.Value_FromNonSelfScript
		} else {
			// Script doesn't reference itself - update both Value and Value_FromNonSelfScript
			cur.Value = newVal
			cur.Value_FromNonSelfScript = newVal
		}
		s.Data[row][col] = cur
		recordChange(row, col, oldVal, cur.Value)
		s.mu.Unlock()
		if triggernext {
			globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
			globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetName, row, col})
			globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
		}

		// Add merged audit entries before save
		addMergedAuditEntries(s, cellChanges)
		globalSheetManager.SaveSheet(s)
		//broadcastRowColUpdated(s, projectName, sheetName)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
		return
	}

	// baseIdx and baseRow already computed above

	// Try to parse as single value (string, number, boolean)
	// If it's not a JSON array or matrix, just set the value to the base cell
	var testInterface interface{}
	if err := json.Unmarshal([]byte(newVal), &testInterface); err != nil {
		// Try to normalize Python-like list syntax to JSON
		tryVal := normalizePythonListToJSON(newVal)
		if err2 := json.Unmarshal([]byte(tryVal), &testInterface); err2 != nil {
			// Not parseable, treat as plain string
			oldVal := previousValues[row+"-"+col]
			if isSelfReferencing {
				cur.Value = cur.Value_FromNonSelfScript
			} else {
				cur.Value = newVal
				cur.Value_FromNonSelfScript = newVal
			}
			s.Data[row][col] = cur
			recordChange(row, col, oldVal, cur.Value)
			s.mu.Unlock()
			addMergedAuditEntries(s, cellChanges)
			globalSheetManager.SaveSheet(s)
			//broadcastRowColUpdated(s, projectName, sheetName)
			globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
			return
		}
		// Use normalized value for subsequent parsing
		newVal = tryVal
	}

	// Check if it's a simple value (not array or object)
	switch testInterface.(type) {
	case string, float64, bool, nil:
		// Simple value, set to base cell
		oldVal := previousValues[row+"-"+col]
		if isSelfReferencing {
			cur.Value = cur.Value_FromNonSelfScript
		} else {
			cur.Value = newVal
			cur.Value_FromNonSelfScript = newVal
		}
		s.Data[row][col] = cur
		recordChange(row, col, oldVal, cur.Value)
		s.mu.Unlock()
		if triggernext {
			globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
			globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{projectName, sheetName, row, col})
			globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
		}

		addMergedAuditEntries(s, cellChanges)
		globalSheetManager.SaveSheet(s)
		//broadcastRowColUpdated(s, projectName, sheetName)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
		return
	} // handling matrix type [[1,2],[3,4]]
	var matrix [][]string
	MatrixParsed := false
	if rSpan > 1 || cSpan > 1 {
		var any interface{}
		tryVal := newVal
		if err := json.Unmarshal([]byte(tryVal), &any); err != nil {
			tryVal = normalizePythonListToJSON(tryVal)
			if err2 := json.Unmarshal([]byte(tryVal), &any); err2 == nil {
				newVal = tryVal
			}
		} else {
			newVal = tryVal
		}
		if any != nil {
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
	CellsModified := make([]CellIdentifier, 0)
	if MatrixParsed {
		// if matrix dimensions is more than ScriptOutput_RowSpan x ScriptOutput_ColSpan, we will simply fill cur.Value = scriptOutput
		if len(matrix) > rSpan || (len(matrix) > 0 && len(matrix[0]) > cSpan) {
			oldVal := previousValues[row+"-"+col]
			if isSelfReferencing {
				cur.Value = cur.Value_FromNonSelfScript
			} else {
				cur.Value = cur.ScriptOutput
				cur.Value_FromNonSelfScript = cur.ScriptOutput
			}
			s.Data[row][col] = cur
			recordChange(row, col, oldVal, cur.Value)
			s.mu.Unlock()
			addMergedAuditEntries(s, cellChanges)
			globalSheetManager.SaveSheet(s)
			//broadcastRowColUpdated(s, projectName, sheetName)
			globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
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
				oldVal := previousValues[rKey+"-"+cLabel]
				// For matrix output, cells other than the script cell are always non-self-referencing
				if dr == 0 && dc == 0 && isSelfReferencing {
					c.Value = c.Value_FromNonSelfScript
				} else {
					c.Value = val
					if dr == 0 && dc == 0 {
						c.Value_FromNonSelfScript = val
					}
				}
				s.Data[rKey][cLabel] = c
				recordChange(rKey, cLabel, oldVal, c.Value)
				if triggernext {
					CellsModified = append(CellsModified, CellIdentifier{projectName, sheetName, rKey, cLabel})
				}
			}
		}

		s.mu.Unlock()
		globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
		globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellsModified...)
		globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
		addMergedAuditEntries(s, cellChanges)
		globalSheetManager.SaveSheet(s)
		//broadcastRowColUpdated(s, projectName, sheetName)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
		return
	}

	// handling array type [1,2,3,4]
	// If output is a flat array and span is 1xN or Nx1, fill accordingly
	var arrAny []interface{}
	tryArr := newVal
	if err := json.Unmarshal([]byte(tryArr), &arrAny); err != nil {
		tryArr = normalizePythonListToJSON(tryArr)
		if err2 := json.Unmarshal([]byte(tryArr), &arrAny); err2 != nil {
			// fall through to plain string handling below
		}
	}
	if arrAny != nil {
		if (rSpan == 1 && cSpan > 1 && len(arrAny) <= cSpan) || (cSpan == 1 && rSpan > 1 && len(arrAny) <= rSpan) {
			for i, v := range arrAny {
				if rSpan == 1 {
					// Fill row left to right
					cLabel := indexToColLabel(baseIdx + i)
					c := s.Data[row][cLabel]
					oldVal := previousValues[row+"-"+cLabel]
					valStr := fmt.Sprint(v)
					c.Value = valStr
					s.Data[row][cLabel] = c
					recordChange(row, cLabel, oldVal, valStr)
					if triggernext {
						CellsModified = append(CellsModified, CellIdentifier{projectName, sheetName, row, cLabel})
					}

				} else {
					// Fill column top to bottom
					rKey := itoa(baseRow + i)
					c := s.Data[rKey][col]
					oldVal := previousValues[rKey+"-"+col]
					valStr := fmt.Sprint(v)
					c.Value = valStr
					s.Data[rKey][col] = c
					recordChange(rKey, col, oldVal, valStr)
					if triggernext {
						CellsModified = append(CellsModified, CellIdentifier{projectName, sheetName, rKey, col})
					}
				}
			}
		} else {
			oldVal := previousValues[row+"-"+col]
			if isSelfReferencing {
				cur.Value = cur.Value_FromNonSelfScript
			} else {
				cur.Value = cur.ScriptOutput
				cur.Value_FromNonSelfScript = cur.ScriptOutput
			}
			s.Data[row][col] = cur
			recordChange(row, col, oldVal, cur.Value)
			if triggernext {
				CellsModified = append(CellsModified, CellIdentifier{projectName, sheetName, row, col})
			}
		}
	}
	s.mu.Unlock()
	globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
	globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellsModified...)
	globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
	addMergedAuditEntries(s, cellChanges)
	globalSheetManager.SaveSheet(s)

	//broadcastRowColUpdated(s, projectName, sheetName)
	globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
}

// ExecuteDependentScripts executes all scripts that depend on the given cell
// This should be called when a cell value is modified to trigger cascading updates
// skipSelf prevent execution of the script in the same cell
func ExecuteDependentScripts(projectName, sheetName, row, col string) {
	// Get the cell ID for the modified cell
	s := globalSheetManager.GetSheetBy(sheetName, projectName)
	if s == nil {
		return
	}
	//fmt.Println("Checking dependents for cell ", projectName, "/", sheetName, " cell ", row+col)
	// Get all dependent scripts
	dependents := globalSheetManager.GetDependentScripts(projectName, sheetName, row, col)

	// Execute each dependent script
	for _, dep := range dependents {
		//fmt.Println("Executing dependent script at ", dep.ScriptProjectName, "/", dep.ScriptSheetName, " cell ", dep.ScriptCellID, " which depends on ", projectName, "/", sheetName, " cell ", row+col)
		ExecuteCellScriptWithIdentifier(dep)
	}
}

func ExecuteCellScriptWithIdentifier(dep ScriptIdentifier) {

	// Find the cell with this script
	depSheet := globalSheetManager.GetSheetBy(dep.ScriptSheetName, dep.ScriptProjectName)
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
		ExecuteCellScript(dep.ScriptProjectName, dep.ScriptSheetName, depRow, depCol)
	}

}

// Removes zombie script locks
func removeLocksWithMissingCellID(s *Sheet) {
	s.mu.Lock()
	for rKey, rowMap := range s.Data {
		for cKey, cell := range rowMap {
			if cell.Locked {
				//identify if the lock is a script-span lock which should have a CellID reference in the LockedBy field like "script-span <CellID>"
				if strings.HasPrefix(cell.LockedBy, "script-span ") {
					// Extract the script cell ID from "script-span <CellID>"
					scriptCellID := strings.TrimPrefix(cell.LockedBy, "script-span ")
					cellpresent := false
					//check if the cell with this CellID exists in the sheet
					for _, rowMap2 := range s.Data {
						for _, cell2 := range rowMap2 {

							if cell2.CellID == scriptCellID {
								cellpresent = true
								break
							}
						}
						if cellpresent {
							break
						}
					}
					if !cellpresent { // If the cell with this CellID exists, do not remove the lock
						cell.Locked = false
						cell.LockedBy = ""
						s.Data[rKey][cKey] = cell
					}
					// Cell exists, do not remove the lockcontinue } } // Cell does not exist, remove the lock cell.Locked = false cell.LockedBy = "" s.Data[rKey][cKey] = cell } else if strings.HasPrefix(cell.LockedBy, "script-") { // If it starts with "script-" but not "script-span", we can also remove the lock cell.Locked = false cell.LockedBy = "" s.Data[rKey][cKey] = cell} else if strings.HasPrefix(cell.LockedBy, "script-") { // If it starts with "script-" but not "script-span", we can also remove the lock cell.Locked = false cell.LockedBy = "" s.Data[rKey][cKey] = cell

				}
			}
		}
	}
	s.mu.Unlock()
	globalSheetManager.SaveSheet(s)
}
