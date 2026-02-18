package main

import (
	"encoding/json"
	"fmt"
	"log"
	"regexp"
	"strings"
	"sync"
	"time"

	"python-libs/data"

	"github.com/kluctl/go-embed-python/embed_util"
	"github.com/kluctl/go-embed-python/python"
)

func init() {
	// Initialize embedded Python once at startup
	if _, err := getEmbeddedPython(); err != nil {
		log.Printf("Embedded Python init failed: %v", err)
	}
}

var (
	embeddedPy        *python.EmbeddedPython
	embeddedPyOnce    sync.Once
	embeddedPyInitErr error
)

// getEmbeddedPython initializes and returns a shared embedded Python interpreter.
// It extracts the embedded Python distribution to a temp dir on first use.
func getEmbeddedPython() (*python.EmbeddedPython, error) {
	var initErr error
	embeddedPyOnce.Do(func() {
		ep, err := python.NewEmbeddedPython("shared-spreadsheet")
		if err != nil {
			initErr = err
			embeddedPyInitErr = err
			return
		}

		libFiles, err := embed_util.NewEmbeddedFiles(data.Data, "python-lib")

		if err != nil {
			initErr = err
			embeddedPyInitErr = err
			return
		}
		fmt.Println("Extracting embedded Python libraries to:", libFiles.GetExtractedPath())
		ep.AddPythonPath(libFiles.GetExtractedPath())
		embeddedPy = ep
	})
	if initErr != nil {
		embeddedPyInitErr = initErr
		return nil, initErr
	}
	return embeddedPy, nil
}

func (sm *SheetManager) initAsyncSaver() {
	sm.mu.Lock()
	defer sm.mu.Unlock()
	if sm.started {
		return
	}
	if sm.saveInterval == 0 {
		sm.saveInterval = 5 * time.Second // default debounce window
	}
	sm.started = true
	go sm.flusher()
}

func (sm *SheetManager) flusher() {
	ticker := time.NewTicker(250 * time.Millisecond)
	defer ticker.Stop()
	for {
		select {
		case <-ticker.C:
			now := time.Now()

			sm.CellsModifiedByScriptQueueMu.Lock()

			if len(sm.CellsModifiedByScriptQueue) == 0 {
				sm.CellsModifiedByScriptQueueMu.Unlock()

				// Process script execution queue
				sm.CellsModifiedManuallyQueueMu.Lock()
				if len(sm.CellsModifiedManuallyQueue) > 0 {
					//clearing ScriptsExecuted list
					sm.ScriptsExecutedMu.Lock()
					//fmt.Println("scripts executed(flusher)", sm.ScriptsExecuted)
					clear(sm.ScriptsExecuted)
					//fmt.Println("scripts executed after clear(flusher)", sm.ScriptsExecuted)
					sm.ScriptsExecutedMu.Unlock()
					toExec := sm.CellsModifiedManuallyQueue[0]
					//pop from queue
					sm.CellsModifiedManuallyQueue = sm.CellsModifiedManuallyQueue[1:]
					sm.CellsModifiedManuallyQueueMu.Unlock()
					//fmt.Println("Executing scripts for manually modified cell:", toExec)
					ExecuteDependentScripts(toExec.ProjectName, toExec.sheetName, toExec.row, toExec.col)
					// Update options for cells which have options depending on range in the sheet of modified cell
					updateOptionsForDependentCells(toExec.ProjectName, toExec.sheetName, toExec.row, toExec.col)

					continue
				} else {
					sm.CellsModifiedManuallyQueueMu.Unlock()
				}
			} else {
				toExec := sm.CellsModifiedByScriptQueue[0]
				//pop from queue
				sm.CellsModifiedByScriptQueue = sm.CellsModifiedByScriptQueue[1:]
				sm.CellsModifiedByScriptQueueMu.Unlock()
				//fmt.Println("Executing scripts for script modified cell:", toExec)
				ExecuteDependentScripts(toExec.ProjectName, toExec.sheetName, toExec.row, toExec.col)
				// Update options for cells which have options depending on range in the sheet of modified cell
				updateOptionsForDependentCells(toExec.ProjectName, toExec.sheetName, toExec.row, toExec.col)

				continue
			}

			// collect due items without holding lock during disk IO
			var toFlush []*Sheet
			sm.mu.Lock()
			for k, ps := range sm.pending {
				if now.Sub(ps.lastModified) >= sm.saveInterval {
					toFlush = append(toFlush, ps.sheet)
					delete(sm.pending, k)
				}
			}
			sm.mu.Unlock()
			// flush outside of lock
			if len(toFlush) > 0 {
				sm.mu.RLock()
				for _, s := range toFlush {
					sm.saveSheetLocked(s)
				}
				sm.mu.RUnlock()
			}

			// Process ROW_COL_UPDATE broadcast queue
			sm.RowColUpdateQueueMu.Lock()
			if len(sm.RowColUpdateQueue) > 0 {
				// Deduplicate sheets in queue - only keep unique project/sheet combinations
				seenSheets := make(map[string]bool)
				uniqueUpdates := make([]RowColUpdateItem, 0)

				for _, item := range sm.RowColUpdateQueue {
					key := item.ProjectName + "::" + item.SheetName
					if !seenSheets[key] {
						seenSheets[key] = true
						uniqueUpdates = append(uniqueUpdates, item)
					}
				}

				// Clear the queue
				sm.RowColUpdateQueue = []RowColUpdateItem{}
				sm.RowColUpdateQueueMu.Unlock()

				// Send ROW_COL_UPDATED messages for each unique sheet
				if globalHub != nil {
					for _, item := range uniqueUpdates {
						sheet := sm.GetSheetBy(item.SheetName, item.ProjectName)
						if sheet != nil {
							payload, _ := json.Marshal(sheet.SnapshotForClient())
							globalHub.broadcast <- &Message{
								Type:      "ROW_COL_UPDATED",
								SheetName: item.SheetName,
								Project:   item.ProjectName,
								Payload:   payload,
								User:      "system",
							}
						}
					}
				}
			} else {
				sm.RowColUpdateQueueMu.Unlock()
			}

			//Code to debug mutex deadlock issues

			{
				currentTime := time.Now()

				// Check SheetManager.mu
				if sm.mu.TryLock() {
					sm.mu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.mu (main mutex) is currently locked at %v", currentTime)
				}

				// Check SheetManager.scriptDepsMu
				if sm.scriptDepsMu.TryLock() {
					sm.scriptDepsMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.scriptDepsMu is currently locked at %v", currentTime)
				}

				// Check SheetManager.OptionsRangeDepsMu
				if sm.OptionsRangeDepsMu.TryLock() {
					sm.OptionsRangeDepsMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.OptionsRangeDepsMu is currently locked at %v", currentTime)
				}

				// Check SheetManager.CellsModifiedManuallyQueueMu
				if sm.CellsModifiedManuallyQueueMu.TryLock() {
					sm.CellsModifiedManuallyQueueMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.CellsModifiedManuallyQueueMu is currently locked at %v", currentTime)
				}

				// Check SheetManager.CellsModifiedByScriptQueueMu
				if sm.CellsModifiedByScriptQueueMu.TryLock() {
					sm.CellsModifiedByScriptQueueMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.CellsModifiedByScriptQueueMu is currently locked at %v", currentTime)
				}

				// Check SheetManager.ScriptsExecutedMu
				if sm.ScriptsExecutedMu.TryLock() {
					sm.ScriptsExecutedMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.ScriptsExecutedMu is currently locked at %v", currentTime)
				}

				// Check SheetManager.RowColUpdateQueueMu
				if sm.RowColUpdateQueueMu.TryLock() {
					sm.RowColUpdateQueueMu.Unlock()
				} else {
					log.Printf("[MUTEX DEBUG] SheetManager.RowColUpdateQueueMu is currently locked at %v", currentTime)
				}

				// Check all sheet mutexes
				if sm.mu.TryRLock() {
					for _, sheet := range sm.sheets {
						if sheet.mu.TryLock() {
							sheet.mu.Unlock()
						} else {
							log.Printf("[MUTEX DEBUG] Sheet %s (Project: %s) mutex is currently locked at %v",
								sheet.Name, sheet.ProjectName, currentTime)
						}
					}
					sm.mu.RUnlock()
				}
			}

		case <-sm.stopCh:
			return
		}
	}
}

// QueueRowColUpdate adds a sheet to the ROW_COL_UPDATE broadcast queue
func (sm *SheetManager) QueueRowColUpdate(projectName, sheetName string) {
	sm.RowColUpdateQueueMu.Lock()
	defer sm.RowColUpdateQueueMu.Unlock()
	// Check if this sheet is already queued to avoid duplicates
	for _, item := range sm.RowColUpdateQueue {
		if item.ProjectName == projectName && item.SheetName == sheetName {
			return
		}
	}
	sm.RowColUpdateQueue = append(sm.RowColUpdateQueue, RowColUpdateItem{
		ProjectName: projectName,
		SheetName:   sheetName,
	})
}

func (s *Sheet) SetCellScript(row, col, script, user string, reverted bool, rowSpan int, colSpan int) {
	s.mu.Lock()
	// ensure row map
	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}
	currentVal, exists := s.Data[row][col]
	// Prevent edits to locked cells
	if exists && currentVal.Locked {
		s.mu.Unlock()
		return
	}
	// Audit only script change
	if reverted {
		prevNew := currentVal.Script
		r1 := atoiSafe(row)
		for i := len(s.AuditLog) - 1; i >= 0; i-- {
			e := &s.AuditLog[i]
			if e.Action == "EDIT_SCRIPT" && e.Row1 == r1 && e.Col1 == col && e.NewValue == prevNew && !e.ChangeReversed {
				e.ChangeReversed = true
				break
			}
		}
	} else {
		var oldScript string
		if exists {
			oldScript = currentVal.Script
		}

		cellChanges := make(map[string]cellChangesstruct)
		// Append edit entry, merging with the last same-user edit for this cell if present
		r1 := atoiSafe(row)
		cLabel := col
		key := row + "-" + cLabel
		cellChanges[key] = cellChangesstruct{
			rowNum: r1,
			colStr: cLabel,
			oldVal: oldScript,
			newVal: script,
			action: "EDIT_SCRIPT",
			user:   user,
		}
		// Add merged audit entries before save
		addMergedAuditEntries(s, cellChanges)
		/*
			prevIdx := -1
			r1 := atoiSafe(row)
			for i := len(s.AuditLog) - 1; i >= 0; i-- {
				entry := s.AuditLog[i]
				if entry.Action == "EDIT_SCRIPT" && entry.Row1 == r1 && entry.Col1 == col {
					if entry.User == user && !entry.ChangeReversed {
						prevIdx = i
					}
					break
				}
			}
			if prevIdx >= 0 {
				if time.Since(s.AuditLog[prevIdx].Timestamp) < 24*time.Hour {
					oldScript = s.AuditLog[prevIdx].OldValue
					s.AuditLog = append(s.AuditLog[:prevIdx], s.AuditLog[prevIdx+1:]...)
				}
			}
			if oldScript != script {
				s.AuditLog = append(s.AuditLog, AuditEntry{
					Timestamp:      time.Now(),
					User:           user,
					Action:         "EDIT_SCRIPT",
					Row1:           r1,
					Col1:           col,
					OldValue:       oldScript,
					NewValue:       script,
					ChangeReversed: false,
				})
			}
		*/
	}

	// Preserve existing metadata
	updated := currentVal
	updated.Script = script
	updated.User = user
	if !exists || exists && strings.TrimSpace(currentVal.CellID) == "" && strings.TrimSpace(script) != "" {
		updated.CellID = generateID()
		updated.CellType = ScriptCell
	}
	if exists && strings.TrimSpace(script) == "" {
		updated.CellType = ValueCell
	}
	// Normalize spans
	if rowSpan <= 0 {
		rowSpan = 1
	}
	if colSpan <= 0 {
		colSpan = 1
	}

	updated.ScriptOutput_RowSpan = rowSpan
	updated.ScriptOutput_ColSpan = colSpan
	s.Data[row][col] = updated

	// Update script dependencies
	cellID := updated.CellID
	s.mu.Unlock()

	// Update dependency map for this script
	globalSheetManager.UpdateScriptDependencies(s.ProjectName, s.Name, cellID, script, row, col)

	// Done updating script; save and execute
	globalSheetManager.SaveSheet(s)
	ExecuteCellScriptonChange(s.ProjectName, s.Name, row, col)
	//s.FillValueFromScriptOutput(row, col)
}

// adjustScriptTagsOnInsertRow increments row references in script tags for rows at or below insertRow
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnInsertRow(insertRow int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	// Collect modified scripts to update dependencies after releasing sheet lock
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := atoiSafe(submatches[2])

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if row1 >= insertRow && row1 > 0 {
						row1++
					}
					return fmt.Sprintf("{{%s%d}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := atoiSafe(submatches[4])

				if row1 >= insertRow && row1 > 0 {
					row1++
				}
				if row2 >= insertRow && row2 > 0 {
					row2++
				}

				return fmt.Sprintf("{{%s%d:%s%d}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	// Update dependencies outside of sheet lock
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s)
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := atoiSafe(submatches[4])

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if row1 >= insertRow && row1 > 0 {
							row1++
						}
						return fmt.Sprintf("{{%s/%s/%s%d}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := atoiSafe(submatches[6])

					if row1 >= insertRow && row1 > 0 {
						row1++
					}
					if row2 >= insertRow && row2 > 0 {
						row2++
					}

					return fmt.Sprintf("{{%s/%s/%s%d:%s%d}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		// Update dependencies for all modified scripts in this sheet without holding sheet lock
		if modified {
			// Snapshot scripts under read lock
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet)
			// Send ROW_COL_UPDATED message to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// adjustScriptTagsOnDeleteRow decrements row references in script tags for rows strictly above deleted row
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnDeleteRow(deleteRow int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := atoiSafe(submatches[2])

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if row1 == deleteRow {
						// Reference to deleted row becomes invalid - keep as is or mark
						return match
					}
					if row1 > deleteRow {
						row1--
					}
					return fmt.Sprintf("{{%s%d}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := atoiSafe(submatches[4])

				// Handle deleted row in range
				if deleteRow >= row1 && deleteRow <= row2 {
					// Row is within range - shrink the range
					if row1 == row2 {
						// Single row range that got deleted - keep as invalid reference
						return match
					}
					if deleteRow == row1 {
						row1++
					} else if deleteRow == row2 {
						row2--
					}
					// If deleteRow is in the middle, just adjust the end
					if row2 > deleteRow {
						row2--
					}
				} else {
					// Adjust if above deleted row
					if row1 > deleteRow {
						row1--
					}
					if row2 > deleteRow {
						row2--
					}
				}

				return fmt.Sprintf("{{%s%d:%s%d}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s)
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := atoiSafe(submatches[4])

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if row1 == deleteRow {
							// Reference to deleted row becomes invalid - keep as is
							return match
						}
						if row1 > deleteRow {
							row1--
						}
						return fmt.Sprintf("{{%s/%s/%s%d}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := atoiSafe(submatches[6])

					// Handle deleted row in range
					if deleteRow >= row1 && deleteRow <= row2 {
						// Row is within range - shrink the range
						if row1 == row2 {
							// Single row range that got deleted - keep as invalid reference
							return match
						}
						if deleteRow == row1 {
							row1++
						} else if deleteRow == row2 {
							row2--
						}
						// If deleteRow is in the middle, just adjust the end
						if row2 > deleteRow {
							row2--
						}
					} else {
						// Adjust if above deleted row
						if row1 > deleteRow {
							row1--
						}
						if row2 > deleteRow {
							row2--
						}
					}

					return fmt.Sprintf("{{%s/%s/%s%d:%s%d}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet)
			// Send ROW_COL_UPDATED m
			// essage to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// adjustScriptTagsOnMoveRow adjusts row references in script tags for a row move from fromRow to destIndex
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnMoveRow(fromRow, destIndex int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := atoiSafe(submatches[2])

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if row1 == 0 {
						return match
					}

					if fromRow < destIndex {
						if row1 == fromRow {
							row1 = destIndex
						} else if row1 > fromRow && row1 <= destIndex {
							row1--
						}
					} else if fromRow > destIndex {
						if row1 == fromRow {
							row1 = destIndex
						} else if row1 >= destIndex && row1 < fromRow {
							row1++
						}
					}
					return fmt.Sprintf("{{%s%d}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := atoiSafe(submatches[4])

				if row1 == 0 || row2 == 0 {
					return match
				}

				if fromRow < destIndex {
					if row1 == fromRow {
						row1 = destIndex
					} else if row1 > fromRow && row1 <= destIndex {
						row1--
					}
					if row2 == fromRow {
						row2 = destIndex
					} else if row2 > fromRow && row2 <= destIndex {
						row2--
					}
				} else if fromRow > destIndex {
					if row1 == fromRow {
						row1 = destIndex
					} else if row1 >= destIndex && row1 < fromRow {
						row1++
					}
					if row2 == fromRow {
						row2 = destIndex
					} else if row2 >= destIndex && row2 < fromRow {
						row2++
					}
				}

				return fmt.Sprintf("{{%s%d:%s%d}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s) // persist the row move before clients fetch updated sheet
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s)
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := atoiSafe(submatches[4])

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if row1 == 0 {
							return match
						}

						if fromRow < destIndex {
							if row1 == fromRow {
								row1 = destIndex
							} else if row1 > fromRow && row1 <= destIndex {
								row1--
							}
						} else if fromRow > destIndex {
							if row1 == fromRow {
								row1 = destIndex
							} else if row1 >= destIndex && row1 < fromRow {
								row1++
							}
						}
						return fmt.Sprintf("{{%s/%s/%s%d}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := atoiSafe(submatches[6])

					if row1 == 0 || row2 == 0 {
						return match
					}

					if fromRow < destIndex {
						if row1 == fromRow {
							row1 = destIndex
						} else if row1 > fromRow && row1 <= destIndex {
							row1--
						}
						if row2 == fromRow {
							row2 = destIndex
						} else if row2 > fromRow && row2 <= destIndex {
							row2--
						}
					} else if fromRow > destIndex {
						if row1 == fromRow {
							row1 = destIndex
						} else if row1 >= destIndex && row1 < fromRow {
							row1++
						}
						if row2 == fromRow {
							row2 = destIndex
						} else if row2 >= destIndex && row2 < fromRow {
							row2++
						}
					}

					return fmt.Sprintf("{{%s/%s/%s%d:%s%d}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet) // persist the row move before clients fetch updated sheet
			// Send ROW_COL_UPDATED message to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// adjustScriptTagsOnInsertCol increments column references in script tags for columns at or beyond insertIdx
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnInsertCol(insertIdx int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := submatches[2]
				col1Idx := colLabelToIndex(col1)

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if col1Idx >= insertIdx && col1Idx > 0 {
						col1Idx++
						col1 = indexToColLabel(col1Idx)
					}
					return fmt.Sprintf("{{%s%s}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := submatches[4]
				col2Idx := colLabelToIndex(col2)

				if col1Idx >= insertIdx && col1Idx > 0 {
					col1Idx++
					col1 = indexToColLabel(col1Idx)
				}
				if col2Idx >= insertIdx && col2Idx > 0 {
					col2Idx++
					col2 = indexToColLabel(col2Idx)
				}

				return fmt.Sprintf("{{%s%s:%s%s}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s) // persist the column insert before clients fetch updated sheet
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := submatches[4]
					col1Idx := colLabelToIndex(col1)

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if col1Idx >= insertIdx && col1Idx > 0 {
							col1Idx++
							col1 = indexToColLabel(col1Idx)
						}
						return fmt.Sprintf("{{%s/%s/%s%s}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := submatches[6]
					col2Idx := colLabelToIndex(col2)

					if col1Idx >= insertIdx && col1Idx > 0 {
						col1Idx++
						col1 = indexToColLabel(col1Idx)
					}
					if col2Idx >= insertIdx && col2Idx > 0 {
						col2Idx++
						col2 = indexToColLabel(col2Idx)
					}

					return fmt.Sprintf("{{%s/%s/%s%s:%s%s}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet) // persist the column insert before clients fetch updated sheet
			// Send ROW_COL_UPDATED message to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// adjustScriptTagsOnDeleteCol decrements column references in script tags for columns strictly right of deleted column
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnDeleteCol(deleteIdx int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := submatches[2]
				col1Idx := colLabelToIndex(col1)

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if col1Idx == deleteIdx {
						// Reference to deleted column becomes invalid - keep as is
						return match
					}
					if col1Idx > deleteIdx {
						col1Idx--
						col1 = indexToColLabel(col1Idx)
					}
					return fmt.Sprintf("{{%s%s}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := submatches[4]
				col2Idx := colLabelToIndex(col2)

				// Handle deleted column in range
				if deleteIdx >= col1Idx && deleteIdx <= col2Idx {
					// Column is within range - shrink the range
					if col1Idx == col2Idx {
						// Single column range that got deleted - keep as invalid reference
						return match
					}
					if deleteIdx == col1Idx {
						col1Idx++
						col1 = indexToColLabel(col1Idx)
					} else if deleteIdx == col2Idx {
						col2Idx--
						col2 = indexToColLabel(col2Idx)
					}
					// If deleteIdx is in the middle, just adjust the end
					if col2Idx > deleteIdx {
						col2Idx--
						col2 = indexToColLabel(col2Idx)
					}
				} else {
					// Adjust if to the right of deleted column
					if col1Idx > deleteIdx {
						col1Idx--
						col1 = indexToColLabel(col1Idx)
					}
					if col2Idx > deleteIdx {
						col2Idx--
						col2 = indexToColLabel(col2Idx)
					}
				}

				return fmt.Sprintf("{{%s%s:%s%s}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s) // persist the column delete before clients fetch updated sheet
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := submatches[4]
					col1Idx := colLabelToIndex(col1)

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if col1Idx == deleteIdx {
							// Reference to deleted column becomes invalid - keep as is
							return match
						}
						if col1Idx > deleteIdx {
							col1Idx--
							col1 = indexToColLabel(col1Idx)
						}
						return fmt.Sprintf("{{%s/%s/%s%s}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := submatches[6]
					col2Idx := colLabelToIndex(col2)

					// Handle deleted column in range
					if deleteIdx >= col1Idx && deleteIdx <= col2Idx {
						// Column is within range - shrink the range
						if col1Idx == col2Idx {
							// Single column range that got deleted - keep as invalid reference
							return match
						}
						if deleteIdx == col1Idx {
							col1Idx++
							col1 = indexToColLabel(col1Idx)
						} else if deleteIdx == col2Idx {
							col2Idx--
							col2 = indexToColLabel(col2Idx)
						}
						// If deleteIdx is in the middle, just adjust the end
						if col2Idx > deleteIdx {
							col2Idx--
							col2 = indexToColLabel(col2Idx)
						}
					} else {
						// Adjust if to the right of deleted column
						if col1Idx > deleteIdx {
							col1Idx--
							col1 = indexToColLabel(col1Idx)
						}
						if col2Idx > deleteIdx {
							col2Idx--
							col2 = indexToColLabel(col2Idx)
						}
					}

					return fmt.Sprintf("{{%s/%s/%s%s:%s%s}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet) // persist the column delete before clients fetch updated sheet
			// Send ROW_COL_UPDATED message to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// adjustScriptTagsOnMoveCol adjusts column references in script tags for a column move from fromIdx to destIdx
// This function updates both same-sheet references and cross-sheet references, and updates scriptDeps
func (s *Sheet) adjustScriptTagsOnMoveCol(fromIdx, destIdx int) {
	sameSheetPattern := regexp.MustCompile(`\{\{([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	crossSheetPattern := regexp.MustCompile(`\{\{([^/\{\}]+)/([^/\{\}]+)/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\}\}`)
	s.mu.Lock()
	// Adjust same-sheet references in this sheet
	type depUpd struct{ project, sheet, cellID, script, row, col string }
	pending := make([]depUpd, 0)
	for rowKey, rowMap := range s.Data {
		for colKey, cell := range rowMap {
			if cell.Script == "" {
				continue
			}

			newScript := sameSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
				submatches := sameSheetPattern.FindStringSubmatch(match)
				if len(submatches) < 3 {
					return match
				}

				col1 := submatches[1]
				row1 := submatches[2]
				col1Idx := colLabelToIndex(col1)

				// Single cell reference {{A2}}
				if submatches[3] == "" || submatches[4] == "" {
					if col1Idx == 0 {
						return match
					}

					oldIdx := col1Idx
					if fromIdx < destIdx {
						if col1Idx == fromIdx {
							col1Idx = destIdx
						} else if col1Idx > fromIdx && col1Idx <= destIdx {
							col1Idx--
						}
					} else if fromIdx > destIdx {
						if col1Idx == fromIdx {
							col1Idx = destIdx
						} else if col1Idx >= destIdx && col1Idx < fromIdx {
							col1Idx++
						}
					}

					if col1Idx != oldIdx {
						col1 = indexToColLabel(col1Idx)
					}
					return fmt.Sprintf("{{%s%s}}", col1, row1)
				}

				// Range reference {{A2:B3}}
				col2 := submatches[3]
				row2 := submatches[4]
				col2Idx := colLabelToIndex(col2)

				if col1Idx == 0 || col2Idx == 0 {
					return match
				}

				oldIdx1 := col1Idx
				oldIdx2 := col2Idx

				if fromIdx < destIdx {
					if col1Idx == fromIdx {
						col1Idx = destIdx
					} else if col1Idx > fromIdx && col1Idx <= destIdx {
						col1Idx--
					}
					if col2Idx == fromIdx {
						col2Idx = destIdx
					} else if col2Idx > fromIdx && col2Idx <= destIdx {
						col2Idx--
					}
				} else if fromIdx > destIdx {
					if col1Idx == fromIdx {
						col1Idx = destIdx
					} else if col1Idx >= destIdx && col1Idx < fromIdx {
						col1Idx++
					}
					if col2Idx == fromIdx {
						col2Idx = destIdx
					} else if col2Idx >= destIdx && col2Idx < fromIdx {
						col2Idx++
					}
				}

				if col1Idx != oldIdx1 {
					col1 = indexToColLabel(col1Idx)
				}
				if col2Idx != oldIdx2 {
					col2 = indexToColLabel(col2Idx)
				}

				return fmt.Sprintf("{{%s%s:%s%s}}", col1, row1, col2, row2)
			})

			if newScript != cell.Script {
				cell.Script = newScript
				s.Data[rowKey][colKey] = cell
				pending = append(pending, depUpd{s.ProjectName, s.Name, cell.CellID, newScript, rowKey, colKey})
			}
		}
	}
	s.mu.Unlock()
	for _, u := range pending {
		globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
	}
	if len(pending) > 0 {
		globalSheetManager.SaveSheet(s) // persist the column move before clients fetch updated sheet
	}
	// Send ROW_COL_UPDATED message to clients if any scripts were modified in this sheet
	if len(pending) > 0 && globalHub != nil {
		globalSheetManager.QueueRowColUpdate(s.ProjectName, s.Name)
	}
	// Adjust cross-sheet references in sheets that reference this sheet
	// Use scriptDeps to find which sheets have dependencies on this sheet
	sheet_Key := s.ProjectName + "/" + s.Name
	globalSheetManager.scriptDepsMu.RLock()
	scriptIdentifiers, hasRefs := globalSheetManager.scriptDeps[sheet_Key]
	globalSheetManager.scriptDepsMu.RUnlock()

	if !hasRefs {
		return // No cross-sheet references to this sheet
	}

	// Collect unique sheets that reference this sheet
	seenSheets := make(map[string]bool)
	sheetsToUpdate := make([]*Sheet, 0)
	globalSheetManager.mu.RLock()
	for _, si := range scriptIdentifiers {
		key := sheetKey(si.ScriptProjectName, si.ScriptSheetName)
		if !seenSheets[key] {
			if sheet, ok := globalSheetManager.sheets[key]; ok && sheet != nil {
				sheetsToUpdate = append(sheetsToUpdate, sheet)
				seenSheets[key] = true
			}
		}
	}
	globalSheetManager.mu.RUnlock()

	for _, sheet := range sheetsToUpdate {
		sheet.mu.Lock()
		modified := false
		for rowKey, rowMap := range sheet.Data {
			for colKey, cell := range rowMap {
				if cell.Script == "" {
					continue
				}

				newScript := crossSheetPattern.ReplaceAllStringFunc(cell.Script, func(match string) string {
					submatches := crossSheetPattern.FindStringSubmatch(match)
					if len(submatches) < 5 {
						return match
					}

					refProject := submatches[1]
					refSheet := submatches[2]

					// Only adjust if referencing the current sheet
					if refProject != s.ProjectName || refSheet != s.Name {
						return match
					}

					col1 := submatches[3]
					row1 := submatches[4]
					col1Idx := colLabelToIndex(col1)

					// Single cell reference {{project/sheet/A2}}
					if submatches[5] == "" || submatches[6] == "" {
						if col1Idx == 0 {
							return match
						}

						oldIdx := col1Idx
						if fromIdx < destIdx {
							if col1Idx == fromIdx {
								col1Idx = destIdx
							} else if col1Idx > fromIdx && col1Idx <= destIdx {
								col1Idx--
							}
						} else if fromIdx > destIdx {
							if col1Idx == fromIdx {
								col1Idx = destIdx
							} else if col1Idx >= destIdx && col1Idx < fromIdx {
								col1Idx++
							}
						}

						if col1Idx != oldIdx {
							col1 = indexToColLabel(col1Idx)
						}
						return fmt.Sprintf("{{%s/%s/%s%s}}", refProject, refSheet, col1, row1)
					}

					// Range reference {{project/sheet/A2:B3}}
					col2 := submatches[5]
					row2 := submatches[6]
					col2Idx := colLabelToIndex(col2)

					if col1Idx == 0 || col2Idx == 0 {
						return match
					}

					oldIdx1 := col1Idx
					oldIdx2 := col2Idx

					if fromIdx < destIdx {
						if col1Idx == fromIdx {
							col1Idx = destIdx
						} else if col1Idx > fromIdx && col1Idx <= destIdx {
							col1Idx--
						}
						if col2Idx == fromIdx {
							col2Idx = destIdx
						} else if col2Idx > fromIdx && col2Idx <= destIdx {
							col2Idx--
						}
					} else if fromIdx > destIdx {
						if col1Idx == fromIdx {
							col1Idx = destIdx
						} else if col1Idx >= destIdx && col1Idx < fromIdx {
							col1Idx++
						}
						if col2Idx == fromIdx {
							col2Idx = destIdx
						} else if col2Idx >= destIdx && col2Idx < fromIdx {
							col2Idx++
						}
					}

					if col1Idx != oldIdx1 {
						col1 = indexToColLabel(col1Idx)
					}
					if col2Idx != oldIdx2 {
						col2 = indexToColLabel(col2Idx)
					}

					return fmt.Sprintf("{{%s/%s/%s%s:%s%s}}", refProject, refSheet, col1, row1, col2, row2)
				})

				if newScript != cell.Script {
					cell.Script = newScript
					sheet.Data[rowKey][colKey] = cell
					modified = true
				}
			}
		}
		sheet.mu.Unlock()

		if modified {
			type depUpd2 struct{ project, sheet, cellID, script, row, col string }
			updList := make([]depUpd2, 0)
			sheet.mu.RLock()
			for rKey, rowMap := range sheet.Data {
				for cKey, cell := range rowMap {
					if cell.Script != "" {
						updList = append(updList, depUpd2{sheet.ProjectName, sheet.Name, cell.CellID, cell.Script, rKey, cKey})
					}
				}
			}
			sheet.mu.RUnlock()
			for _, u := range updList {
				globalSheetManager.UpdateScriptDependencies(u.project, u.sheet, u.cellID, u.script, u.row, u.col)
			}
			globalSheetManager.SaveSheet(sheet) // persist the column move before clients fetch updated sheet
			// Send ROW_COL_UPDATED message to clients for this modified sheet
			if globalHub != nil {
				globalSheetManager.QueueRowColUpdate(sheet.ProjectName, sheet.Name)
			}
		}
	}
}

// rebuildScriptDependencies rebuilds the script dependency map from all loaded sheets
// This should be called after loading sheets from disk on startup
func (sm *SheetManager) rebuildScriptDependencies() {
	sm.scriptDepsMu.Lock()
	defer sm.scriptDepsMu.Unlock()

	// Clear existing dependencies
	sm.scriptDeps = make(map[string][]ScriptIdentifier)

	// Iterate through all sheets and extract dependencies from scripts
	for _, sheet := range sm.sheets {
		if sheet == nil {
			continue
		}

		sheet.mu.RLock()
		projectName := sheet.ProjectName
		sheetName := sheet.Name

		// Iterate with keys to capture row and column labels
		for rowLabel, rowMap := range sheet.Data {
			for colLabel, cell := range rowMap {
				if strings.TrimSpace(cell.Script) == "" {
					continue
				}

				// Extract dependencies for this script
				deps := ExtractScriptDependencies(cell.Script, projectName, sheetName)
				if len(deps) == 0 {
					// No explicit references found; add self cell as dependency
					deps = append(deps, DependencyInfo{
						Project: projectName,
						Sheet:   sheetName,
						Range:   colLabel + rowLabel,
					})
				}
				// Add to dependency map
				for _, dep := range deps {
					sheetKey := dep.Project + "/" + dep.Sheet
					scriptIdent := ScriptIdentifier{
						ScriptProjectName: projectName,
						ScriptSheetName:   sheetName,
						ScriptCellID:      cell.CellID,
						ReferencedRange:   dep.Range,
					}
					sm.scriptDeps[sheetKey] = append(sm.scriptDeps[sheetKey], scriptIdent)
				}
			}
		}
		sheet.mu.RUnlock()
	}

	log.Printf("Rebuilt script dependency map with %d referenced sheets", len(sm.scriptDeps))
}
