package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

func (s *Sheet) extractOptionsFromRange(rangeStr string) []string {
	rangeStr = strings.TrimSpace(rangeStr)
	if rangeStr == "" {
		return nil
	}

	// Check if it's a cross-sheet reference (project/.../sheetid/A1:A10).
	// The range portion is always the last slash-segment; the sheet name is the second-to-last;
	// everything before that forms the project path (may contain slashes for subfolders).
	var targetSheet *Sheet
	var rangeOnly string

	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) >= 3 {
		// Cross-sheet reference: project/.../sheetid/A1:A10
		n := len(slashParts)
		rangeOnly = slashParts[n-1]
		refSheetName := slashParts[n-2]
		refProjectName := strings.Join(slashParts[:n-2], "/")

		// Get the referenced sheet
		targetSheet = globalSheetManager.GetSheetBy(refSheetName, refProjectName)
		if targetSheet == nil {
			return nil
		}
	} else {
		// Same-sheet reference: A1:A10
		targetSheet = s
		rangeOnly = rangeStr
	}

	// Parse range format: "A1:B10" or "A1:A10"
	parts := strings.Split(rangeOnly, ":")
	if len(parts) != 2 {
		return nil
	}

	// Parse start cell (e.g., "A1")
	startCell := strings.TrimSpace(parts[0])
	endCell := strings.TrimSpace(parts[1])

	// Extract column and row from start cell
	startCol := ""
	startRow := ""
	for i, ch := range startCell {
		if ch >= '0' && ch <= '9' {
			startCol = strings.ToUpper(startCell[:i])
			startRow = startCell[i:]
			break
		}
	}

	// Extract column and row from end cell
	endCol := ""
	endRow := ""
	for i, ch := range endCell {
		if ch >= '0' && ch <= '9' {
			endCol = strings.ToUpper(endCell[:i])
			endRow = endCell[i:]
			break
		}
	}

	if startCol == "" || startRow == "" || endCol == "" || endRow == "" {
		return nil
	}

	// Convert to indices
	startColIdx := colLabelToIndex(startCol)
	endColIdx := colLabelToIndex(endCol)
	startRowNum := atoiSafe(startRow)
	endRowNum := atoiSafe(endRow)

	if startColIdx == 0 || endColIdx == 0 || startRowNum == 0 || endRowNum == 0 {
		return nil
	}

	// Ensure start <= end
	if startColIdx > endColIdx {
		startColIdx, endColIdx = endColIdx, startColIdx
	}
	if startRowNum > endRowNum {
		startRowNum, endRowNum = endRowNum, startRowNum
	}

	// Lock the target sheet for reading
	targetSheet.mu.RLock()
	defer targetSheet.mu.RUnlock()

	// Extract values from the range only if it's a single row or single column
	var options []string
	isSingleRow := startRowNum == endRowNum
	isSingleColumn := startColIdx == endColIdx

	if isSingleRow || isSingleColumn {
		for rowNum := startRowNum; rowNum <= endRowNum; rowNum++ {
			rowStr := strconv.Itoa(rowNum)
			for colIdx := startColIdx; colIdx <= endColIdx; colIdx++ {
				colStr := indexToColLabel(colIdx)
				if targetSheet.Data[rowStr] != nil {
					if cell, exists := targetSheet.Data[rowStr][colStr]; exists {
						value := strings.TrimSpace(cell.Value)
						if value != "" {
							options = append(options, value)
						}
					}
				}
			}
		}
	}

	return options
}

// SetCellType updates cell type for a cell. Only owner can change cell type.
func (s *Sheet) SetCellType(row, col string, cellType int, options []string, optionsRange string, user string) bool {
	s.mu.Lock()
	//defer s.mu.Unlock()

	// Only owner can change cell type
	if user != s.Owner {
		s.mu.Unlock()
		return false
	}

	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}

	current := s.Data[row][col]
	oldOptions := append([]string(nil), current.Options...)
	oldSelected := append([]int(nil), current.OptionsSelected...)
	s.mu.Unlock()
	// If optionsRange is provided, extract options from the specified range
	if optionsRange != "" {
		extractedOptions := s.extractOptionsFromRange(optionsRange)
		if len(extractedOptions) > 0 {
			options = extractedOptions
		}
	}
	s.mu.Lock()
	// Update cell type and options
	current.CellType = cellType
	current.Options = options
	current.OptionsRange = optionsRange
	current.User = user

	// Clear selected options when changing cell type
	if cellType != ComboBoxCell && cellType != MultipleSelectionCell {
		current.OptionsSelected = nil
	}

	// If options changed, update Value based on previous selection
	optionsChanged := len(oldOptions) != len(options)
	if !optionsChanged {
		for i := range oldOptions {
			if oldOptions[i] != options[i] {
				optionsChanged = true
				break
			}
		}
	}

	if optionsChanged && (cellType == ComboBoxCell || cellType == MultipleSelectionCell) {
		if cellType == ComboBoxCell {
			if len(oldSelected) > 0 {
				idx := oldSelected[0]
				if idx >= 0 && idx < len(options) {
					current.Value = options[idx]
					current.OptionsSelected = []int{idx}
				} else {
					current.Value = ""
					current.OptionsSelected = nil
				}
			} else {
				current.Value = ""
				current.OptionsSelected = nil
			}
		} else if cellType == MultipleSelectionCell {
			var selectedValues []string
			var validIdx []int
			for _, idx := range oldSelected {
				if idx >= 0 && idx < len(options) {
					selectedValues = append(selectedValues, options[idx])
					validIdx = append(validIdx, idx)
				}
			}
			current.Value = strings.Join(selectedValues, "; ")
			current.OptionsSelected = validIdx
		}
	}

	s.Data[row][col] = current

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp:      time.Now(),
		User:           user,
		Action:         "CHANGE_CELL_TYPE",
		Row1:           atoiSafe(row),
		Col1:           col,
		ChangeReversed: false,
	})
	s.mu.Unlock()

	// Update OptionsRange dependencies
	globalSheetManager.UpdateOptionsRangeDependencies(s.ProjectName, s.Name, row, col, optionsRange)

	globalSheetManager.SaveSheet(s)
	return true
}

// SetCellOptionSelected updates the selected options for a ComboBox or MultipleSelection cell
func (s *Sheet) SetCellOptionSelected(row, col string, optionSelected []int) {
	s.mu.Lock()
	defer s.mu.Unlock()

	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}

	current := s.Data[row][col]
	current.OptionsSelected = optionSelected
	s.Data[row][col] = current
}

func (sm *SheetManager) rebuildOptionsRangeDependencies() {
	sm.OptionsRangeDepsMu.Lock()
	defer sm.OptionsRangeDepsMu.Unlock()

	// Clear existing dependencies
	sm.OptionsRangeDeps = make(map[string][]CellIdentifier)

	// Iterate through all sheets and extract OptionsRange dependencies
	for _, sheet := range sm.sheets {
		if sheet == nil {
			continue
		}

		sheet.mu.RLock()
		projectName := sheet.ProjectName
		sheetName := sheet.Name

		for rowLabel, rowMap := range sheet.Data {
			for colLabel, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				// Parse the OptionsRange to determine which sheet it depends on
				deps := parseOptionsRangeDependency(cell.OptionsRange, projectName, sheetName)
				for _, dep := range deps {
					depKey := dep.Project + "/" + dep.Sheet
					sm.OptionsRangeDeps[depKey] = append(sm.OptionsRangeDeps[depKey], CellIdentifier{
						ProjectName: projectName,
						sheetName:   sheetName,
						row:         rowLabel,
						col:         colLabel,
					})
				}
			}
		}
		sheet.mu.RUnlock()
	}

	log.Printf("Rebuilt OptionsRange dependency map with %d referenced sheets", len(sm.OptionsRangeDeps))
}

// parseOptionsRangeDependency parses an OptionsRange string and returns dependency info.
// OptionsRange can be "A1:A10" (same sheet) or "projectname/sheetid/A1:A10" (cross-sheet).
func parseOptionsRangeDependency(optionsRange, currentProject, currentSheet string) []DependencyInfo {
	optionsRange = strings.TrimSpace(optionsRange)
	if optionsRange == "" {
		return nil
	}

	var deps []DependencyInfo

	slashParts := strings.Split(optionsRange, "/")
	if len(slashParts) >= 3 {
		// Cross-sheet reference: project/.../sheetid/A1:A10
		// The range is the last segment, sheet name is second-to-last,
		// and the project path (may contain slashes for subfolders) is everything before.
		n := len(slashParts)
		deps = append(deps, DependencyInfo{
			Project: strings.Join(slashParts[:n-2], "/"),
			Sheet:   slashParts[n-2],
			Range:   slashParts[n-1],
		})
	} else {
		// Same-sheet reference: A1:A10
		deps = append(deps, DependencyInfo{
			Project: currentProject,
			Sheet:   currentSheet,
			Range:   optionsRange,
		})
	}

	return deps
}

// UpdateOptionsRangeDependencies updates the OptionsRange dependency map for a cell.
// Should be called whenever a cell's OptionsRange is set or changed.
func (sm *SheetManager) UpdateOptionsRangeDependencies(cellProjectName, cellSheetName, cellRow, cellCol, optionsRange string) {
	sm.OptionsRangeDepsMu.Lock()
	defer sm.OptionsRangeDepsMu.Unlock()

	// Remove old dependencies for this cell (check all sheet keys)
	for depKey, cells := range sm.OptionsRangeDeps {
		filtered := make([]CellIdentifier, 0, len(cells))
		for _, c := range cells {
			if c.ProjectName != cellProjectName || c.sheetName != cellSheetName || c.row != cellRow || c.col != cellCol {
				filtered = append(filtered, c)
			}
		}
		if len(filtered) > 0 {
			sm.OptionsRangeDeps[depKey] = filtered
		} else {
			delete(sm.OptionsRangeDeps, depKey)
		}
	}

	// If optionsRange is empty, we're done (dependencies removed)
	if strings.TrimSpace(optionsRange) == "" {
		return
	}

	// Add new dependencies
	deps := parseOptionsRangeDependency(optionsRange, cellProjectName, cellSheetName)
	for _, dep := range deps {
		depKey := dep.Project + "/" + dep.Sheet
		sm.OptionsRangeDeps[depKey] = append(sm.OptionsRangeDeps[depKey], CellIdentifier{
			ProjectName: cellProjectName,
			sheetName:   cellSheetName,
			row:         cellRow,
			col:         cellCol,
		})
	}
}

// RenameProjectInOptionsRangeDependencies updates OptionsRange dependency references when a project is renamed.
// It also updates the OptionsRange field in all cells that reference the old project name.
func (sm *SheetManager) RenameProjectInOptionsRangeDependencies(oldProject, newProject string) {
	sm.OptionsRangeDepsMu.Lock()
	defer sm.OptionsRangeDepsMu.Unlock()

	// Create new map with updated keys
	newDeps := make(map[string][]CellIdentifier)

	for depKey, cells := range sm.OptionsRangeDeps {
		newDepKey := depKey
		parts := strings.Split(depKey, "/")
		if len(parts) >= 2 && parts[0] == oldProject {
			parts[0] = newProject
			newDepKey = strings.Join(parts, "/")
		}

		// Update cell identifiers (including those in subfolders)
		newCells := make([]CellIdentifier, 0, len(cells))
		for _, c := range cells {
			if c.ProjectName == oldProject {
				c.ProjectName = newProject
			} else if strings.HasPrefix(c.ProjectName, oldProject+"/") {
				c.ProjectName = newProject + c.ProjectName[len(oldProject):]
			}
			newCells = append(newCells, c)
		}

		newDeps[newDepKey] = newCells
	}

	sm.OptionsRangeDeps = newDeps

	// Update OptionsRange field in all cells that reference the old project name
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
				if strings.TrimSpace(cell.OptionsRange) != "" {
					oldPattern := oldProject + "/"
					if strings.HasPrefix(cell.OptionsRange, oldPattern) {
						cell.OptionsRange = newProject + "/" + cell.OptionsRange[len(oldPattern):]
						sheet.Data[rowKey][colKey] = cell
						modified = true
					}
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sm.SaveSheet(sheet)
		}
	}
}

// RenameSheetInOptionsRangeDependencies updates OptionsRange dependency references when a sheet is renamed within a project.
// It also updates the OptionsRange field in all cells that reference the old sheet name.
func (sm *SheetManager) RenameSheetInOptionsRangeDependencies(projectName, oldSheetName, newSheetName string) {
	sm.OptionsRangeDepsMu.Lock()
	defer sm.OptionsRangeDepsMu.Unlock()

	// Create new map with updated keys
	newDeps := make(map[string][]CellIdentifier)

	for depKey, cells := range sm.OptionsRangeDeps {
		newDepKey := depKey
		parts := strings.Split(depKey, "/")
		// depKey format is "project/sheet"
		if len(parts) >= 2 && parts[0] == projectName && parts[1] == oldSheetName {
			parts[1] = newSheetName
			newDepKey = strings.Join(parts, "/")
		}

		// Update cell identifiers that belong to the renamed sheet
		newCells := make([]CellIdentifier, 0, len(cells))
		for _, c := range cells {
			if c.ProjectName == projectName && c.sheetName == oldSheetName {
				c.sheetName = newSheetName
			}
			newCells = append(newCells, c)
		}

		newDeps[newDepKey] = newCells
	}

	sm.OptionsRangeDeps = newDeps

	// Update OptionsRange field in all cells that reference the old sheet name
	// Pattern: "projectname/oldSheetName/A1:A10" -> "projectname/newSheetName/A1:A10"
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
				if strings.TrimSpace(cell.OptionsRange) != "" {
					oldPattern := projectName + "/" + oldSheetName + "/"
					if strings.HasPrefix(cell.OptionsRange, oldPattern) {
						cell.OptionsRange = projectName + "/" + newSheetName + "/" + cell.OptionsRange[len(oldPattern):]
						sheet.Data[rowKey][colKey] = cell
						modified = true
					}
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sm.SaveSheet(sheet)
		}
	}
}

// adjustOptionsRangeOnInsertRow adjusts OptionsRange references in all cells when a row is inserted.
// This handles same-sheet references inline and cross-sheet references via OptionsRangeDeps.
func (s *Sheet) adjustOptionsRangeOnInsertRow(insertRow int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnInsertRow(cell.OptionsRange, insertRow, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				// Re-extract options from the new range
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		// Re-extract options for modified cells and update deps
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	// Adjust cross-sheet OptionsRange references
	s.adjustCrossSheetOptionsRangeOnInsertRow(insertRow)
}

// adjustOptionsRangeOnDeleteRow adjusts OptionsRange references in all cells when a row is deleted.
func (s *Sheet) adjustOptionsRangeOnDeleteRow(deleteRow int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnDeleteRow(cell.OptionsRange, deleteRow, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnDeleteRow(deleteRow)
}

// adjustOptionsRangeOnMoveRow adjusts OptionsRange references in all cells when a row is moved.
func (s *Sheet) adjustOptionsRangeOnMoveRow(fromRow, destIndex int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnMoveRow(cell.OptionsRange, fromRow, destIndex, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnMoveRow(fromRow, destIndex)
}

// adjustOptionsRangeOnMoveRowBlock adjusts OptionsRange references for a block move.
// blockStart is the original first row, blockSize is the number of rows, insertStart is the destination.
func (s *Sheet) adjustOptionsRangeOnMoveRowBlock(blockStart, blockSize, insertStart int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnMoveRowBlock(cell.OptionsRange, blockStart, blockSize, insertStart, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnMoveRowBlock(blockStart, blockSize, insertStart)
}

// adjustOptionsRangeOnDeleteRowBlock adjusts OptionsRange references for a block delete.
// blockStart is the first row deleted, blockSize is the number of contiguous rows deleted.
func (s *Sheet) adjustOptionsRangeOnDeleteRowBlock(blockStart, blockSize int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnDeleteRowBlock(cell.OptionsRange, blockStart, blockSize, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnDeleteRowBlock(blockStart, blockSize)
}

// adjustOptionsRangeOnInsertCol adjusts OptionsRange references in all cells when a column is inserted.
func (s *Sheet) adjustOptionsRangeOnInsertCol(insertIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnInsertCol(cell.OptionsRange, insertIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnInsertCol(insertIdx)
}

// adjustOptionsRangeOnDeleteCol adjusts OptionsRange references in all cells when a column is deleted.
func (s *Sheet) adjustOptionsRangeOnDeleteCol(deleteIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnDeleteCol(cell.OptionsRange, deleteIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnDeleteCol(deleteIdx)
}

// adjustOptionsRangeOnMoveCol adjusts OptionsRange references in all cells when a column is moved.
func (s *Sheet) adjustOptionsRangeOnMoveCol(fromIdx, destIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	s.mu.Lock()
	modified := false
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) == "" {
				continue
			}
			newRange := adjustRangeRefOnMoveCol(cell.OptionsRange, fromIdx, destIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
			if newRange != cell.OptionsRange {
				cell.OptionsRange = newRange
				s.Data[rowKey][colKey] = cell
				modified = true
			}
		}
	}
	s.mu.Unlock()

	if modified {
		s.refreshOptionsFromRanges()
		globalSheetManager.SaveSheet(s)
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}

	s.adjustCrossSheetOptionsRangeOnMoveCol(fromIdx, destIdx)
}

// refreshOptionsFromRanges re-extracts options from OptionsRange for all cells that have one set.
// Should be called after adjusting OptionsRange references.
func (s *Sheet) refreshOptionsFromRanges() {
	s.mu.Lock()
	// Collect cells that need options refresh
	type cellRef struct {
		row, col     string
		optionsRange string
	}
	var toRefresh []cellRef
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if strings.TrimSpace(cell.OptionsRange) != "" {
				toRefresh = append(toRefresh, cellRef{rowKey, colKey, cell.OptionsRange})
			}
		}
	}
	s.mu.Unlock()

	for _, ref := range toRefresh {
		extractedOptions := s.extractOptionsFromRange(ref.optionsRange)
		s.mu.Lock()
		if len(extractedOptions) > 0 {
			cell := s.Data[ref.row][ref.col]
			cell.Options = extractedOptions
			s.Data[ref.row][ref.col] = cell
		}
		s.mu.Unlock()
	}

	// Update OptionsRange dependencies for all modified cells
	for _, ref := range toRefresh {
		globalSheetManager.UpdateOptionsRangeDependencies(s.ProjectName, s.Name, ref.row, ref.col, ref.optionsRange)
	}
}

// Helper functions for adjusting range references

func adjustRangeRefOnInsertRow(rangeStr string, insertRow int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		// Cross-sheet: project/sheet/A1:B10
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr // Not referencing the sheet being modified
		}
		newRange := adjustSameSheetRangeOnInsertRow(rRange, insertRow, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	// Same-sheet
	return adjustSameSheetRangeOnInsertRow(rangeStr, insertRow, samePattern)
}

func adjustSameSheetRangeOnInsertRow(rangeStr string, insertRow int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := atoiSafe(m[2])
	col2 := m[3]
	row2 := atoiSafe(m[4])

	if row1 >= insertRow && row1 > 0 {
		row1++
	}
	if row2 >= insertRow && row2 > 0 {
		row2++
	}
	return fmt.Sprintf("%s%d:%s%d", col1, row1, col2, row2)
}

func adjustRangeRefOnDeleteRow(rangeStr string, deleteRow int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnDeleteRow(rRange, deleteRow, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnDeleteRow(rangeStr, deleteRow, samePattern)
}

func adjustSameSheetRangeOnDeleteRow(rangeStr string, deleteRow int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := atoiSafe(m[2])
	col2 := m[3]
	row2 := atoiSafe(m[4])

	if deleteRow >= row1 && deleteRow <= row2 {
		if row1 == row2 {
			return rangeStr // Single row range deleted - keep as is
		}
		row2--
	} else {
		if row1 > deleteRow {
			row1--
		}
		if row2 > deleteRow {
			row2--
		}
	}
	return fmt.Sprintf("%s%d:%s%d", col1, row1, col2, row2)
}

func adjustRangeRefOnMoveRow(rangeStr string, fromRow, destIndex int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnMoveRow(rRange, fromRow, destIndex, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnMoveRow(rangeStr, fromRow, destIndex, samePattern)
}

func adjustSameSheetRangeOnMoveRow(rangeStr string, fromRow, destIndex int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := atoiSafe(m[2])
	col2 := m[3]
	row2 := atoiSafe(m[4])

	adjustRow := func(r int) int {
		if fromRow < destIndex {
			if r == fromRow {
				return destIndex
			} else if r > fromRow && r <= destIndex {
				return r - 1
			}
		} else if fromRow > destIndex {
			if r == fromRow {
				return destIndex
			} else if r >= destIndex && r < fromRow {
				return r + 1
			}
		}
		return r
	}

	row1 = adjustRow(row1)
	row2 = adjustRow(row2)
	return fmt.Sprintf("%s%d:%s%d", col1, row1, col2, row2)
}

func adjustRangeRefOnMoveRowBlock(rangeStr string, blockStart, blockSize, insertStart int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnMoveRowBlock(rRange, blockStart, blockSize, insertStart, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnMoveRowBlock(rangeStr, blockStart, blockSize, insertStart, samePattern)
}

func adjustSameSheetRangeOnMoveRowBlock(rangeStr string, blockStart, blockSize, insertStart int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := atoiSafe(m[2])
	col2 := m[3]
	row2 := atoiSafe(m[4])

	adjustRow := func(r int) int {
		blockEnd := blockStart + blockSize - 1
		if r >= blockStart && r <= blockEnd {
			return insertStart + (r - blockStart)
		}
		if blockStart < insertStart {
			if r > blockEnd && r < insertStart+blockSize {
				return r - blockSize
			}
		} else if blockStart > insertStart {
			if r >= insertStart && r < blockStart {
				return r + blockSize
			}
		}
		return r
	}

	row1 = adjustRow(row1)
	row2 = adjustRow(row2)
	return fmt.Sprintf("%s%d:%s%d", col1, row1, col2, row2)
}

func adjustRangeRefOnDeleteRowBlock(rangeStr string, blockStart, blockSize int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnDeleteRowBlock(rRange, blockStart, blockSize, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnDeleteRowBlock(rangeStr, blockStart, blockSize, samePattern)
}

func adjustSameSheetRangeOnDeleteRowBlock(rangeStr string, blockStart, blockSize int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := atoiSafe(m[2])
	col2 := m[3]
	row2 := atoiSafe(m[4])
	blockEnd := blockStart + blockSize - 1

	if blockStart <= row1 && blockEnd >= row2 {
		// Entire range deleted - keep as is
		return rangeStr
	}
	if blockStart >= row1 && blockEnd <= row2 {
		// Block is within range - shrink
		row2 -= blockSize
	} else if blockStart <= row1 && blockEnd >= row1 && blockEnd < row2 {
		// Block overlaps start of range
		row1 = blockStart
		row2 -= blockSize
	} else if blockStart > row1 && blockStart <= row2 && blockEnd >= row2 {
		// Block overlaps end of range
		row2 = blockStart - 1
	} else {
		if row1 > blockEnd {
			row1 -= blockSize
		}
		if row2 > blockEnd {
			row2 -= blockSize
		}
	}
	return fmt.Sprintf("%s%d:%s%d", col1, row1, col2, row2)
}

func adjustRangeRefOnInsertCol(rangeStr string, insertIdx int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnInsertCol(rRange, insertIdx, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnInsertCol(rangeStr, insertIdx, samePattern)
}

func adjustSameSheetRangeOnInsertCol(rangeStr string, insertIdx int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := m[2]
	col2 := m[3]
	row2 := m[4]

	col1Idx := colLabelToIndex(col1)
	col2Idx := colLabelToIndex(col2)

	if col1Idx >= insertIdx && col1Idx > 0 {
		col1Idx++
		col1 = indexToColLabel(col1Idx)
	}
	if col2Idx >= insertIdx && col2Idx > 0 {
		col2Idx++
		col2 = indexToColLabel(col2Idx)
	}
	return fmt.Sprintf("%s%s:%s%s", col1, row1, col2, row2)
}

func adjustRangeRefOnDeleteCol(rangeStr string, deleteIdx int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnDeleteCol(rRange, deleteIdx, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnDeleteCol(rangeStr, deleteIdx, samePattern)
}

func adjustSameSheetRangeOnDeleteCol(rangeStr string, deleteIdx int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := m[2]
	col2 := m[3]
	row2 := m[4]

	col1Idx := colLabelToIndex(col1)
	col2Idx := colLabelToIndex(col2)

	if deleteIdx >= col1Idx && deleteIdx <= col2Idx {
		if col1Idx == col2Idx {
			return rangeStr // Single column range deleted - keep as is
		}
		col2Idx--
		col2 = indexToColLabel(col2Idx)
	} else {
		if col1Idx > deleteIdx {
			col1Idx--
			col1 = indexToColLabel(col1Idx)
		}
		if col2Idx > deleteIdx {
			col2Idx--
			col2 = indexToColLabel(col2Idx)
		}
	}
	return fmt.Sprintf("%s%s:%s%s", col1, row1, col2, row2)
}

func adjustRangeRefOnMoveCol(rangeStr string, fromIdx, destIdx int, refProject, refSheet, curProject, curSheet string, samePattern, crossPattern *regexp.Regexp) string {
	slashParts := strings.Split(rangeStr, "/")
	if len(slashParts) == 3 {
		rProject := slashParts[0]
		rSheet := slashParts[1]
		rRange := slashParts[2]
		if rProject != curProject || rSheet != curSheet {
			return rangeStr
		}
		newRange := adjustSameSheetRangeOnMoveCol(rRange, fromIdx, destIdx, samePattern)
		return rProject + "/" + rSheet + "/" + newRange
	}
	return adjustSameSheetRangeOnMoveCol(rangeStr, fromIdx, destIdx, samePattern)
}

func adjustSameSheetRangeOnMoveCol(rangeStr string, fromIdx, destIdx int, pattern *regexp.Regexp) string {
	m := pattern.FindStringSubmatch(rangeStr)
	if m == nil {
		return rangeStr
	}
	col1 := m[1]
	row1 := m[2]
	col2 := m[3]
	row2 := m[4]

	col1Idx := colLabelToIndex(col1)
	col2Idx := colLabelToIndex(col2)

	adjustCol := func(idx int) int {
		if fromIdx < destIdx {
			if idx == fromIdx {
				return destIdx
			} else if idx > fromIdx && idx <= destIdx {
				return idx - 1
			}
		} else if fromIdx > destIdx {
			if idx == fromIdx {
				return destIdx
			} else if idx >= destIdx && idx < fromIdx {
				return idx + 1
			}
		}
		return idx
	}

	newCol1Idx := adjustCol(col1Idx)
	newCol2Idx := adjustCol(col2Idx)

	if newCol1Idx != col1Idx {
		col1 = indexToColLabel(newCol1Idx)
	}
	if newCol2Idx != col2Idx {
		col2 = indexToColLabel(newCol2Idx)
	}
	return fmt.Sprintf("%s%s:%s%s", col1, row1, col2, row2)
}

// Cross-sheet OptionsRange adjustment helpers

func (s *Sheet) adjustCrossSheetOptionsRangeOnInsertRow(insertRow int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	// Group by sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue // Same sheet - already handled
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnInsertRow(cell.OptionsRange, insertRow, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnDeleteRow(deleteRow int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnDeleteRow(cell.OptionsRange, deleteRow, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnMoveRow(fromRow, destIndex int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnMoveRow(cell.OptionsRange, fromRow, destIndex, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnMoveRowBlock(blockStart, blockSize, insertStart int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnMoveRowBlock(cell.OptionsRange, blockStart, blockSize, insertStart, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnDeleteRowBlock(blockStart, blockSize int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnDeleteRowBlock(cell.OptionsRange, blockStart, blockSize, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnInsertCol(insertIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnInsertCol(cell.OptionsRange, insertIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnDeleteCol(deleteIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnDeleteCol(cell.OptionsRange, deleteIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

func (s *Sheet) adjustCrossSheetOptionsRangeOnMoveCol(fromIdx, destIdx int) {
	sameSheetRangePattern := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	crossSheetRangePattern := regexp.MustCompile(`^([^/]+)/([^/]+)/([A-Z]+)(\d+):([A-Z]+)(\d+)$`)

	sheetKey := s.ProjectName + "/" + s.Name
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()
	if !hasDeps {
		return
	}

	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if dep.ProjectName == s.ProjectName && dep.sheetName == s.Name {
			continue
		}
		if !seenSheets[sk] {
			sheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[sk] = true
			}
		}
	}

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if strings.TrimSpace(cell.OptionsRange) == "" {
					continue
				}
				newRange := adjustRangeRefOnMoveCol(cell.OptionsRange, fromIdx, destIdx, s.ProjectName, s.Name, s.ProjectName, s.Name, sameSheetRangePattern, crossSheetRangePattern)
				if newRange != cell.OptionsRange {
					cell.OptionsRange = newRange
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			sheet.refreshOptionsFromRanges()
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}

// CopyPasteProject copies all sheets from sourcePath to destPath, rewriting cross-sheet
// references from the source prefix to the destination prefix.
// sourcePath can be a top-level project or a subfolder (e.g. "proj/sub1").
// destPath is the full destination path (e.g. "proj2" or "proj/sub2").
// The newOwner will be set as owner for all copied sheets (and ensured in editors).
// Returns error if destPath is inside sourcePath (pasting into itself).
func (sm *SheetManager) CopyPasteProject(sourcePath, destPath, newOwner string) error {
	if sourcePath == "" || destPath == "" {
		return fmt.Errorf("source and destination paths required")
	}

	// Prevent pasting inside itself: destPath must not equal sourcePath or be a child of it
	if destPath == sourcePath || strings.HasPrefix(destPath, sourcePath+"/") {
		return fmt.Errorf("cannot paste a folder inside itself")
	}

	// Ensure top-level destination directory
	if err := os.MkdirAll(filepath.Join(dataDir, destPath), 0755); err != nil {
		return fmt.Errorf("failed to create dest directory: %w", err)
	}

	// rewriteRef replaces every occurrence of the sourcePath prefix in a cross-sheet
	// reference string (Script or OptionsRange) with destPath.
	// Cross-sheet refs are written as "sourcePath/..." so a simple prefix replacement is safe.
	rewriteRef := func(ref string) string {
		oldPrefix := sourcePath + "/"
		newPrefix := destPath + "/"
		// Also handle exact match (e.g. "sourcePath/sheetId/A1")
		ref = strings.ReplaceAll(ref, oldPrefix, newPrefix)
		return ref
	}

	sm.mu.Lock()
	defer sm.mu.Unlock()

	for _, s := range sm.sheets {
		if s == nil {
			continue
		}
		// Include sheets at the source level AND all subfolder sheets
		if s.ProjectName != sourcePath && !strings.HasPrefix(s.ProjectName, sourcePath+"/") {
			continue
		}

		// Compute the ProjectName for the clone: replace the leading sourcePath with destPath
		var cloneProjectName string
		if s.ProjectName == sourcePath {
			cloneProjectName = destPath
		} else {
			cloneProjectName = destPath + s.ProjectName[len(sourcePath):]
		}

		// Ensure the subfolder directory exists in the new tree
		if err := os.MkdirAll(filepath.Join(dataDir, cloneProjectName), 0755); err != nil {
			return fmt.Errorf("failed to create subfolder %s: %w", cloneProjectName, err)
		}

		// Build permissions for the clone
		perms := s.Permissions
		hasOwner := false
		for _, e := range perms.Editors {
			if e == newOwner {
				hasOwner = true
				break
			}
		}
		if !hasOwner && newOwner != "" {
			perms.Editors = append(perms.Editors, newOwner)
		}

		clone := &Sheet{
			Name:        s.Name,
			Owner:       newOwner,
			ProjectName: cloneProjectName,
			Data:        make(map[string]map[string]Cell),
			ColWidths:   make(map[string]int),
			RowHeights:  make(map[string]int),
			Permissions: perms,
			AuditLog:    append([]AuditEntry{}, s.AuditLog...),
		}

		// Deep copy cell data, rewriting any cross-sheet references from source to dest
		s.mu.RLock()
		for r, cols := range s.Data {
			clone.Data[r] = make(map[string]Cell, len(cols))
			for c, cell := range cols {
				// Rewrite Script: replace {{sourcePath/...}} with {{destPath/...}}
				if strings.TrimSpace(cell.Script) != "" {
					cell.Script = rewriteRef(cell.Script)
				}
				// Rewrite OptionsRange: replace sourcePath/... with destPath/...
				if strings.TrimSpace(cell.OptionsRange) != "" {
					cell.OptionsRange = rewriteRef(cell.OptionsRange)
				}
				clone.Data[r][c] = cell
			}
		}
		for k, v := range s.ColWidths {
			clone.ColWidths[k] = v
		}
		for k, v := range s.RowHeights {
			clone.RowHeights[k] = v
		}
		s.mu.RUnlock()

		// Register in memory and persist to disk
		sm.sheets[sheetKey(cloneProjectName, clone.Name)] = clone
		sm.saveSheetLocked(clone)
	}
	return nil
}

// updateOptionsForDependentCells updates options for combo box/multiple selection cells
// that depend on the modified cell's sheet via OptionsRange
func updateOptionsForDependentCells(projectName, sheetName, row, col string) {
	sheetKey := projectName + "/" + sheetName

	// Get cells that depend on this sheet for their options
	globalSheetManager.OptionsRangeDepsMu.RLock()
	deps, hasDeps := globalSheetManager.OptionsRangeDeps[sheetKey]
	globalSheetManager.OptionsRangeDepsMu.RUnlock()

	if !hasDeps {
		return
	}

	// Group dependent cells by sheet to minimize sheet lookups
	type sheetCells struct {
		sheet *Sheet
		cells []CellIdentifier
	}

	sheetMap := make(map[string]*sheetCells)

	for _, dep := range deps {
		sk := dep.ProjectName + "::" + dep.sheetName
		if sheetMap[sk] == nil {
			depSheet := globalSheetManager.GetSheetBy(dep.sheetName, dep.ProjectName)
			if depSheet == nil {
				continue
			}
			sheetMap[sk] = &sheetCells{
				sheet: depSheet,
				cells: []CellIdentifier{},
			}
		}
		sheetMap[sk].cells = append(sheetMap[sk].cells, dep)
	}

	// Update each dependent sheet
	for _, sc := range sheetMap {
		sheet := sc.sheet
		modified := false

		for _, dep := range sc.cells {
			sheet.mu.Lock()
			if sheet.Data[dep.row] == nil || sheet.Data[dep.row][dep.col].OptionsRange == "" {
				sheet.mu.Unlock()
				continue
			}

			cell := sheet.Data[dep.row][dep.col]

			// Skip if cell is not a combo box or multiple selection
			if cell.CellType != ComboBoxCell && cell.CellType != MultipleSelectionCell {
				sheet.mu.Unlock()
				continue
			}

			optionsRange := cell.OptionsRange
			sheet.mu.Unlock()

			// Extract new options from the range
			extractedOptions := sheet.extractOptionsFromRange(optionsRange)

			if len(extractedOptions) > 0 {
				depRowInt := -1
				if dr, err := strconv.Atoi(dep.row); err == nil {
					depRowInt = dr
				}
				if sheet.SheetType == "document" && depRowInt == 2 {
					// Document sheet: propagate new options to all  the rows in the same column on the basis options in row 2.

					for targetRow := range sheet.Data {
						if targetRow == "1" {
							continue
						}
						sheet.mu.Lock()
						if sheet.Data[targetRow] == nil {
							sheet.mu.Unlock()
							continue
						}
						row2Cell := sheet.Data["2"][dep.col]
						targetCell := sheet.Data[targetRow][dep.col]
						oldOptions := targetCell.Options
						targetCell.Options = extractedOptions

						// Update value based on OptionsSelected
						if row2Cell.CellType == ComboBoxCell {
							if len(targetCell.OptionsSelected) > 0 && targetCell.OptionsSelected[0] < len(row2Cell.Options) {
								targetCell.Value = row2Cell.Options[targetCell.OptionsSelected[0]]
							} else {
								targetCell.Value = ""
								targetCell.OptionsSelected = nil
							}
						} else if row2Cell.CellType == MultipleSelectionCell {
							var selectedValues []string
							validIndices := []int{}
							for _, idx := range targetCell.OptionsSelected {
								if idx < len(row2Cell.Options) {
									selectedValues = append(selectedValues, row2Cell.Options[idx])
									validIndices = append(validIndices, idx)
								}
							}
							targetCell.Value = strings.Join(selectedValues, "; ")
							targetCell.OptionsSelected = validIndices
						}

						sheet.Data[targetRow][dep.col] = targetCell
						sheet.mu.Unlock()

						modified = true

						optionsChanged := len(oldOptions) != len(extractedOptions)
						if !optionsChanged {
							for i := range oldOptions {
								if oldOptions[i] != extractedOptions[i] {
									optionsChanged = true
									break
								}
							}
						}
						if optionsChanged {
							globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
							globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{dep.ProjectName, dep.sheetName, targetRow, dep.col})
							globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
						}
					}
				} else {
					// Original behaviour for non-document sheets (or document rows other than 2).
					sheet.mu.Lock()
					cell = sheet.Data[dep.row][dep.col]
					oldOptions := cell.Options
					cell.Options = extractedOptions

					// Update value based on OptionsSelected
					if cell.CellType == ComboBoxCell {
						// For combo box, set value from the selected option
						if len(cell.OptionsSelected) > 0 && cell.OptionsSelected[0] < len(cell.Options) {
							cell.Value = cell.Options[cell.OptionsSelected[0]]
						} else {
							// If OptionsSelected is invalid or empty, clear the value
							cell.Value = ""
							cell.OptionsSelected = nil
						}
					} else if cell.CellType == MultipleSelectionCell {
						// For multiple selection, concatenate selected values with semicolon
						var selectedValues []string
						validIndices := []int{}
						for _, idx := range cell.OptionsSelected {
							if idx < len(cell.Options) {
								selectedValues = append(selectedValues, cell.Options[idx])
								validIndices = append(validIndices, idx)
							}
						}
						cell.Value = strings.Join(selectedValues, "; ")
						cell.OptionsSelected = validIndices
					}

					sheet.Data[dep.row][dep.col] = cell
					sheet.mu.Unlock()

					modified = true

					// Check if options actually changed
					optionsChanged := len(oldOptions) != len(extractedOptions)
					if !optionsChanged {
						for i := range oldOptions {
							if oldOptions[i] != extractedOptions[i] {
								optionsChanged = true
								break
							}
						}
					}

					// If options changed, we may need to re-execute dependent scripts
					if optionsChanged {
						// Queue this cell for script execution if it has dependent scripts
						globalSheetManager.CellsModifiedByScriptQueueMu.Lock()
						globalSheetManager.CellsModifiedByScriptQueue = append(globalSheetManager.CellsModifiedByScriptQueue, CellIdentifier{dep.ProjectName, dep.sheetName, dep.row, dep.col})
						globalSheetManager.CellsModifiedByScriptQueueMu.Unlock()
					}
				}
			}
		}

		if modified {
			globalSheetManager.SaveSheet(sheet)
			globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
		}
	}
}
