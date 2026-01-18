package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sync"
	"time"
)

const dataDir = "DATA"

// getSheetFilePath returns the file path for a sheet
func getSheetFilePath(sheetID string) string {
	return filepath.Join(dataDir, sheetID+".json")
}

// ensureDataDir creates the DATA directory if it doesn't exist
func ensureDataDir() error {
	return os.MkdirAll(dataDir, 0755)
}

type Cell struct {
	Value      string `json:"value"`
	User       string `json:"user,omitempty"` // Last edited by
	Locked     bool   `json:"locked,omitempty"`
	LockedBy   string `json:"locked_by,omitempty"`
	Background string `json:"background,omitempty"`
	Bold       bool   `json:"bold,omitempty"`
	Italic     bool   `json:"italic,omitempty"`
}

type AuditEntry struct {
	Timestamp time.Time `json:"timestamp"`
	User      string    `json:"user"`
	Action    string    `json:"action"` // e.g., "EDIT_CELL", "CREATE_SHEET"
	Details   string    `json:"details"`
}

type Permissions struct {
	Editors []string `json:"editors"`
}

type Sheet struct {
	ID          string                     `json:"id"`
	Name        string                     `json:"name"`
	Owner       string                     `json:"owner"`
	ProjectName string                     `json:"project_name,omitempty"`
	Data        map[string]map[string]Cell `json:"data"` // Row -> Col -> Cell
	AuditLog    []AuditEntry               `json:"audit_log"`
	Permissions Permissions                `json:"permissions"`
	ColWidths   map[string]int             `json:"col_widths,omitempty"`
	RowHeights  map[string]int             `json:"row_heights,omitempty"`
	mu          sync.RWMutex
}

// IsEditor returns true if user is the owner or listed as an editor.
func (s *Sheet) IsEditor(user string) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if user == "" {
		return false
	}
	if user == s.Owner {
		return true
	}
	for _, e := range s.Permissions.Editors {
		if e == user {
			return true
		}
	}
	return false
}

type SheetManager struct {
	// Keyed by composite key of project+id
	sheets map[string]*Sheet
	mu     sync.RWMutex
}

// Helper to save a single sheet without locking the manager (caller must hold lock)
func (sm *SheetManager) saveSheetLocked(sheet *Sheet) {
	if err := ensureDataDir(); err != nil {
		log.Printf("Error creating data directory: %v", err)
		return
	}

	// Determine path based on project folder if present
	var dir string
	if sheet.ProjectName != "" {
		dir = filepath.Join(dataDir, sheet.ProjectName)
		if err := os.MkdirAll(dir, 0755); err != nil {
			log.Printf("Error creating project directory %s: %v", dir, err)
			return
		}
	} else {
		dir = dataDir
	}
	filePath := filepath.Join(dir, sheet.ID+".json")
	file, err := os.Create(filePath)
	if err != nil {
		log.Printf("Error saving sheet %s: %v", sheet.ID, err)
		return
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(sheet); err != nil {
		log.Printf("Error encoding sheet %s: %v", sheet.ID, err)
	}
}

// MarshalJSON implementation for Sheet to ensure thread-safe encoding
func (s *Sheet) MarshalJSON() ([]byte, error) {
	s.mu.RLock()
	defer s.mu.RUnlock()

	type Alias Sheet
	return json.Marshal(&struct {
		*Alias
	}{
		Alias: (*Alias)(s),
	})
}

var globalSheetManager = &SheetManager{
	sheets: make(map[string]*Sheet),
}

// sheetKey builds a unique key combining project and sheet id.
func sheetKey(project, id string) string {
	if project == "" {
		return id
	}
	return project + "::" + id
}

func (sm *SheetManager) CreateSheet(name, owner, projectName string) *Sheet {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	id := generateID() // Need to implement this or use a simple counter
	sheet := &Sheet{
		ID:          id,
		Name:        name,
		Owner:       owner,
		ProjectName: projectName,
		Data:        make(map[string]map[string]Cell),
		ColWidths:   make(map[string]int),
		RowHeights:  make(map[string]int),
		Permissions: Permissions{
			Editors: []string{owner},
		},
		AuditLog: []AuditEntry{},
	}

	// Initial Audit
	sheet.AuditLog = append(sheet.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      owner,
		Action:    "CREATE_SHEET",
		Details:   "Created sheet " + name,
	})

	sm.sheets[sheetKey(projectName, id)] = sheet
	sm.saveSheetLocked(sheet) // Persist individual sheet
	return sheet
}

// GetSheet finds a sheet by id only. If multiple projects contain the same id,
// the returned sheet is undefined. Prefer GetSheetBy.
func (sm *SheetManager) GetSheet(id string) *Sheet {
	sm.mu.RLock()
	defer sm.mu.RUnlock()
	for _, s := range sm.sheets {
		if s != nil && s.ID == id {
			return s
		}
	}
	return nil
}

// GetSheetBy finds a sheet by id and project name.
func (sm *SheetManager) GetSheetBy(id, project string) *Sheet {
	sm.mu.RLock()
	defer sm.mu.RUnlock()
	// Try direct composite key first
	if s, ok := sm.sheets[sheetKey(project, id)]; ok {
		return s
	}
	// Fallback: iterate (handles legacy keys)
	for _, s := range sm.sheets {
		if s != nil && s.ID == id && s.ProjectName == project {
			return s
		}
	}
	return nil
}

// CopySheetToProject creates a copy of source sheet into target project.
// New sheet ID is generated; name defaults to source name if empty.
func (sm *SheetManager) CopySheetToProject(sourceID, sourceProject, targetProject, newName, owner string) *Sheet {
	// Locate source
	src := sm.GetSheetBy(sourceID, sourceProject)
	if src == nil {
		return nil
	}
	// Build new sheet
	sm.mu.Lock()
	defer sm.mu.Unlock()
	id := generateID()
	if newName == "" {
		newName = src.Name
	}
	copySheet := &Sheet{
		ID:          id,
		Name:        newName,
		Owner:       owner,
		ProjectName: targetProject,
		Data:        make(map[string]map[string]Cell),
		ColWidths:   make(map[string]int),
		RowHeights:  make(map[string]int),
		Permissions: Permissions{Editors: []string{owner}},
		AuditLog:    append([]AuditEntry{}, src.AuditLog...),
	}
	// Deep copy data
	src.mu.RLock()
	for r, cols := range src.Data {
		copySheet.Data[r] = make(map[string]Cell, len(cols))
		for c, cell := range cols {
			copySheet.Data[r][c] = cell
		}
	}
	for k, v := range src.ColWidths {
		copySheet.ColWidths[k] = v
	}
	for k, v := range src.RowHeights {
		copySheet.RowHeights[k] = v
	}
	src.mu.RUnlock()
	// Register and persist
	sm.sheets[sheetKey(targetProject, id)] = copySheet
	sm.saveSheetLocked(copySheet)
	return copySheet
}

// Simple ID generator
func generateID() string {
	return time.Now().Format("20060102150405")
}

func (s *Sheet) SetCell(row, col, value, user string) {
	s.mu.Lock()
	// defer s.mu.Unlock() // MOVED to explicit unlock before Save()

	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}
	currentVal, exists := s.Data[row][col]
	// Prevent edits to locked cells
	if exists && currentVal.Locked {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "EDIT_DENIED",
			Details:   "Attempted edit on locked cell " + row + "," + col,
		})
		s.mu.Unlock()
		return
	}
	if exists && currentVal.Value == value {
		// No change
		s.mu.Unlock()
		return
	}
	// Preserve existing lock metadata on write
	s.Data[row][col] = Cell{Value: value, User: user, Locked: currentVal.Locked, LockedBy: currentVal.LockedBy, Background: currentVal.Background, Bold: currentVal.Bold, Italic: currentVal.Italic}
	if exists {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "EDIT_CELL",
			Details:   "Changed cell " + row + "," + col + " from " + currentVal.Value + " to " + value,
		})
	} else {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "EDIT_CELL",
			Details:   "Set cell " + row + "," + col + " to " + value,
		})
	}

	s.mu.Unlock() // Unlock BEFORE saving to avoid deadlock (Save -> MarshalJSON -> tries RLock)

	// Persist changes
	// Optimally we shouldn't save on every cell edit for performance, but for this task it ensures safety.
	globalSheetManager.SaveSheet(s)
}

// SetCellStyle updates style attributes for a cell while preserving value and lock metadata.
func (s *Sheet) SetCellStyle(row, col, background string, bold, italic bool, user string) {
	s.mu.Lock()
	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}
	current, exists := s.Data[row][col]
	// Prevent edits to locked cells' style if locked
	if exists && current.Locked {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "STYLE_DENIED",
			Details:   "Attempted style change on locked cell " + row + "," + col,
		})
		s.mu.Unlock()
		return
	}
	// Apply style while preserving existing value and lock info
	updated := current
	updated.User = user
	updated.Background = background
	updated.Bold = bold
	updated.Italic = italic
	s.Data[row][col] = updated

	if exists {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "STYLE_CELL",
			Details:   "Updated style for cell " + row + "," + col,
		})
	} else {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp: time.Now(),
			User:      user,
			Action:    "STYLE_CELL",
			Details:   "Set style for new cell " + row + "," + col,
		})
	}
	s.mu.Unlock()
	globalSheetManager.SaveSheet(s)
}

// IsCellLocked returns whether the given cell is locked.
func (s *Sheet) IsCellLocked(row, col string) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.Data[row] == nil {
		return false
	}
	c, ok := s.Data[row][col]
	if !ok {
		return false
	}
	return c.Locked
}

// LockCell locks a cell. Only the sheet owner may lock.
func (s *Sheet) LockCell(row, col, user string) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	if user != s.Owner {
		return false
	}
	if s.Data[row] == nil {
		s.Data[row] = make(map[string]Cell)
	}
	cell := s.Data[row][col]
	if cell.Locked {
		return true // already locked
	}
	cell.Locked = true
	cell.LockedBy = user
	s.Data[row][col] = cell
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "LOCK_CELL",
		Details:   "Locked cell " + row + "," + col,
	})
	// Save after unlock via manager
	go globalSheetManager.SaveSheet(s)
	return true
}

// UnlockCell unlocks a cell. Only the sheet owner may unlock.
func (s *Sheet) UnlockCell(row, col, user string) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	if user != s.Owner {
		return false
	}
	cell, ok := s.Data[row][col]
	if !ok {
		return false
	}
	if !cell.Locked {
		return true // already unlocked
	}
	cell.Locked = false
	cell.LockedBy = ""
	s.Data[row][col] = cell
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "UNLOCK_CELL",
		Details:   "Unlocked cell " + row + "," + col,
	})
	// Save after unlock via manager
	go globalSheetManager.SaveSheet(s)
	return true
}

func (s *Sheet) SetColWidth(col string, width int, user string) {
	s.mu.Lock()
	// ensure map
	if s.ColWidths == nil {
		s.ColWidths = make(map[string]int)
	}
	s.ColWidths[col] = width

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "RESIZE_COL",
		Details:   "Set width of column " + col + " to " + itoa(width),
	})
	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
}

func (s *Sheet) SetRowHeight(row string, height int, user string) {
	s.mu.Lock()
	if s.RowHeights == nil {
		s.RowHeights = make(map[string]int)
	}
	s.RowHeights[row] = height

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "RESIZE_ROW",
		Details:   "Set height of row " + row + " to " + itoa(height),
	})
	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
}

// UpdatePermissions sets editors list; only owner may change settings.
// Ensures owner is always in editors.
func (s *Sheet) UpdatePermissions(editors []string, performedBy string) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	if performedBy != s.Owner {
		return false
	}
	// dedupe helpers
	uniq := func(in []string) []string {
		m := make(map[string]struct{})
		out := make([]string, 0, len(in))
		for _, v := range in {
			if v == "" {
				continue
			}
			if _, ok := m[v]; !ok {
				m[v] = struct{}{}
				out = append(out, v)
			}
		}
		return out
	}
	editors = uniq(editors)
	// Ensure owner in editors
	hasOwner := false
	for _, e := range editors {
		if e == s.Owner {
			hasOwner = true
			break
		}
	}
	if !hasOwner {
		editors = append(editors, s.Owner)
	}

	s.Permissions.Editors = editors
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      performedBy,
		Action:    "UPDATE_PERMISSIONS",
		Details:   fmt.Sprintf("Editors: %v", editors),
	})
	go globalSheetManager.SaveSheet(s)
	return true
}

// TransferOwnership changes the owner to newOwner; only current owner may transfer.
// New owner is ensured in editors list.
func (s *Sheet) TransferOwnership(newOwner, performedBy string) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	if performedBy != s.Owner {
		return false
	}
	old := s.Owner
	if newOwner == "" || newOwner == old {
		return false
	}
	s.Owner = newOwner
	// Ensure new owner in editors
	found := false
	for _, e := range s.Permissions.Editors {
		if e == newOwner {
			found = true
			break
		}
	}
	if !found {
		s.Permissions.Editors = append(s.Permissions.Editors, newOwner)
	}
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      performedBy,
		Action:    "TRANSFER_OWNERSHIP",
		Details:   fmt.Sprintf("Owner changed from %s to %s", old, newOwner),
	})
	go globalSheetManager.SaveSheet(s)
	return true
}

// InsertRowBelow inserts a new empty row directly below `targetRowStr`, shifting subsequent rows (data and heights) down by one.
// Returns true if an insertion occurred.
func (s *Sheet) InsertRowBelow(targetRowStr, user string) bool {
	var targetRow int
	if _, err := fmt.Sscanf(targetRowStr, "%d", &targetRow); err != nil {
		return false
	}

	insertRow := targetRow + 1

	s.mu.Lock()

	// Shift existing rows [insertRow..] down by 1
	maxRow := 0
	for rowKey := range s.Data {
		var r int
		if _, err := fmt.Sscanf(rowKey, "%d", &r); err == nil {
			if r > maxRow {
				maxRow = r
			}
		}
	}
	for r := maxRow; r >= insertRow; r-- {
		fromKey := itoa(r)
		toKey := itoa(r + 1)
		if rowData, ok := s.Data[fromKey]; ok {
			delete(s.Data, fromKey)
			s.Data[toKey] = rowData
		} else {
			delete(s.Data, toKey)
		}
	}

	// Ensure the new row exists but empty
	newKey := itoa(insertRow)
	if s.Data == nil {
		s.Data = make(map[string]map[string]Cell)
	}
	if _, ok := s.Data[newKey]; !ok {
		s.Data[newKey] = make(map[string]Cell)
	}

	// Shift RowHeights
	if s.RowHeights == nil {
		s.RowHeights = make(map[string]int)
	}
	maxHeightRow := 0
	for rowKey := range s.RowHeights {
		var r int
		if _, err := fmt.Sscanf(rowKey, "%d", &r); err == nil {
			if r > maxHeightRow {
				maxHeightRow = r
			}
		}
	}
	for r := maxHeightRow; r >= insertRow; r-- {
		fromKey := itoa(r)
		toKey := itoa(r + 1)
		if h, ok := s.RowHeights[fromKey]; ok {
			delete(s.RowHeights, fromKey)
			s.RowHeights[toKey] = h
		} else {
			delete(s.RowHeights, toKey)
		}
	}
	// New row height defaults to existing height of target row, if any
	if h, ok := s.RowHeights[targetRowStr]; ok {
		s.RowHeights[newKey] = h
	}

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "INSERT_ROW",
		Details:   "Inserted row " + newKey + " below row " + targetRowStr,
	})

	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
	return true
}

// MoveRowBelow moves the row `fromRowStr` to be directly below `targetRowStr`.
// It shifts the intervening rows accordingly and preserves cell contents and row heights.
// Returns true if a move occurred.
func (s *Sheet) MoveRowBelow(fromRowStr, targetRowStr, user string) bool {
	// Parse integers
	var fromRow, targetRow int
	if _, err := fmt.Sscanf(fromRowStr, "%d", &fromRow); err != nil {
		return false
	}
	if _, err := fmt.Sscanf(targetRowStr, "%d", &targetRow); err != nil {
		return false
	}

	destIndex := targetRow + 1
	if destIndex == fromRow { // no-op
		return false
	}
	if fromRow < destIndex {
		destIndex-- // Adjust for removal before insertion
	}
	s.mu.Lock()
	// Prevent cutting a row containing locked cells
	if rowMap, ok := s.Data[fromRowStr]; ok {
		for _, cell := range rowMap {
			if cell.Locked {
				s.AuditLog = append(s.AuditLog, AuditEntry{
					Timestamp: time.Now(),
					User:      user,
					Action:    "MOVE_ROW_DENIED",
					Details:   "Attempted cut/move of locked row " + fromRowStr,
				})
				s.mu.Unlock()
				return false
			}
		}
	}
	// Snapshot cells for affected range
	start := fromRow
	end := destIndex
	if start > end {
		start, end = end, start
	}

	cellsByRowBefore := make(map[int]map[string]Cell)
	for r := start; r <= end; r++ {
		rowKey := itoa(r)
		if m, ok := s.Data[rowKey]; ok {
			clone := make(map[string]Cell, len(m))
			for c, cell := range m {
				clone[c] = cell
			}
			cellsByRowBefore[r] = clone
		} else {
			cellsByRowBefore[r] = make(map[string]Cell)
		}
	}

	savedRowCells := cellsByRowBefore[fromRow]

	// Helper to clear a row
	clearRow := func(row int) {
		rowKey := itoa(row)
		delete(s.Data, rowKey)
	}

	// Perform shifts
	if fromRow < destIndex {
		// Move down: shift [fromRow+1..destIndex] up by 1
		for k := fromRow + 1; k <= destIndex; k++ {
			target := k - 1
			clearRow(target)
			fromMap := cellsByRowBefore[k]
			if len(fromMap) > 0 {
				s.Data[itoa(target)] = make(map[string]Cell, len(fromMap))
				for col, cell := range fromMap {
					s.Data[itoa(target)][col] = cell
				}
			}
		}
		// Place saved row at destIndex
		clearRow(destIndex)
		if len(savedRowCells) > 0 {
			s.Data[itoa(destIndex)] = make(map[string]Cell, len(savedRowCells))
			for col, cell := range savedRowCells {
				s.Data[itoa(destIndex)][col] = cell
			}
		}
	} else {
		// Move up: shift [destIndex..fromRow-1] down by 1
		for k := fromRow - 1; k >= destIndex; k-- {
			target := k + 1
			clearRow(target)
			fromMap := cellsByRowBefore[k]
			if len(fromMap) > 0 {
				s.Data[itoa(target)] = make(map[string]Cell, len(fromMap))
				for col, cell := range fromMap {
					s.Data[itoa(target)][col] = cell
				}
			}
		}
		// Place saved row at destIndex
		clearRow(destIndex)
		if len(savedRowCells) > 0 {
			s.Data[itoa(destIndex)] = make(map[string]Cell, len(savedRowCells))
			for col, cell := range savedRowCells {
				s.Data[itoa(destIndex)][col] = cell
			}
		}
	}

	// Update RowHeights
	if s.RowHeights == nil {
		s.RowHeights = make(map[string]int)
	}
	newHeights := make(map[string]int, len(s.RowHeights))
	for k, v := range s.RowHeights {
		newHeights[k] = v
	}
	if fromRow < destIndex {
		for k := fromRow + 1; k <= destIndex; k++ {
			newHeights[itoa(k-1)] = s.RowHeights[itoa(k)]
		}
		newHeights[itoa(destIndex)] = s.RowHeights[itoa(fromRow)]
	} else {
		for k := fromRow - 1; k >= destIndex; k-- {
			newHeights[itoa(k+1)] = s.RowHeights[itoa(k)]
		}
		newHeights[itoa(destIndex)] = s.RowHeights[itoa(fromRow)]
	}
	s.RowHeights = newHeights

	// Audit entry
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "MOVE_ROW",
		Details:   fmt.Sprintf("Moved row %d to below row %d", fromRow, targetRow),
	})

	s.mu.Unlock()

	// Save after unlock
	globalSheetManager.SaveSheet(s)
	return true
}

// MoveColumnRight moves the column `fromColStr` to be directly right of `targetColStr`.
// It shifts the intervening columns accordingly and preserves cell contents and column widths.
// Returns true if a move occurred.
func (s *Sheet) MoveColumnRight(fromColStr, targetColStr, user string) bool {
	// Convert column labels to indices (A, B, C, ...)
	toColIdx := func(label string) int {
		idx := 0
		for i := 0; i < len(label); i++ {
			idx = idx*26 + int(label[i]-'A'+1)
		}
		return idx
	}
	toColLabel := func(idx int) string {
		label := ""
		for idx > 0 {
			idx--
			b := byte(int('A') + (idx % 26))
			label = string([]byte{b}) + label
			idx /= 26
		}
		return label
	}

	fromIdx := toColIdx(fromColStr)
	targetIdx := toColIdx(targetColStr)
	if fromIdx == 0 || targetIdx == 0 {
		return false
	}
	destIdx := targetIdx + 1
	if destIdx == fromIdx {
		return false
	}
	if fromIdx < destIdx {
		destIdx-- // Adjust for removal before insertion
	}

	s.mu.Lock()
	// Prevent cutting a column containing any locked cell
	for rowKey, rowMap := range s.Data {
		if cell, ok := rowMap[fromColStr]; ok {
			if cell.Locked {
				s.AuditLog = append(s.AuditLog, AuditEntry{
					Timestamp: time.Now(),
					User:      user,
					Action:    "MOVE_COL_DENIED",
					Details:   "Attempted cut/move of locked column " + fromColStr + " (row " + rowKey + ")",
				})
				s.mu.Unlock()
				return false
			}
		}
	}
	// Find all affected columns
	start := fromIdx
	end := destIdx
	if start > end {
		start, end = end, start
	}

	// Snapshot cells for affected columns
	cellsByColBefore := make(map[int]map[string]Cell)
	for c := start; c <= end; c++ {
		colLabel := toColLabel(c)
		colCells := make(map[string]Cell)
		for rowKey, rowMap := range s.Data {
			if cell, ok := rowMap[colLabel]; ok {
				colCells[rowKey] = cell
			}
		}
		cellsByColBefore[c] = colCells
	}
	savedColCells := cellsByColBefore[fromIdx]

	// Helper to clear a column
	clearCol := func(colIdx int) {
		colLabel := toColLabel(colIdx)
		for _, rowMap := range s.Data {
			delete(rowMap, colLabel)
		}
	}

	// Perform shifts
	if fromIdx < destIdx {
		// Move right: shift [fromIdx+1..destIdx] left by 1
		for k := fromIdx + 1; k <= destIdx; k++ {
			target := k - 1
			clearCol(target)
			fromMap := cellsByColBefore[k]
			for rowKey, cell := range fromMap {
				if s.Data[rowKey] == nil {
					s.Data[rowKey] = make(map[string]Cell)
				}
				s.Data[rowKey][toColLabel(target)] = cell
			}
		}
		// Place saved col at destIdx
		clearCol(destIdx)
		for rowKey, cell := range savedColCells {
			if s.Data[rowKey] == nil {
				s.Data[rowKey] = make(map[string]Cell)
			}
			s.Data[rowKey][toColLabel(destIdx)] = cell
		}
	} else {
		// Move left: shift [destIdx..fromIdx-1] right by 1
		for k := fromIdx - 1; k >= destIdx; k-- {
			target := k + 1
			clearCol(target)
			fromMap := cellsByColBefore[k]
			for rowKey, cell := range fromMap {
				if s.Data[rowKey] == nil {
					s.Data[rowKey] = make(map[string]Cell)
				}
				s.Data[rowKey][toColLabel(target)] = cell
			}
		}
		// Place saved col at destIdx
		clearCol(destIdx)
		for rowKey, cell := range savedColCells {
			if s.Data[rowKey] == nil {
				s.Data[rowKey] = make(map[string]Cell)
			}
			s.Data[rowKey][toColLabel(destIdx)] = cell
		}
	}

	// Update ColWidths
	if s.ColWidths == nil {
		s.ColWidths = make(map[string]int)
	}
	newWidths := make(map[string]int, len(s.ColWidths))
	for k, v := range s.ColWidths {
		newWidths[k] = v
	}
	if fromIdx < destIdx {
		for k := fromIdx + 1; k <= destIdx; k++ {
			newWidths[toColLabel(k-1)] = s.ColWidths[toColLabel(k)]
		}
		newWidths[toColLabel(destIdx)] = s.ColWidths[toColLabel(fromIdx)]
	} else {
		for k := fromIdx - 1; k >= destIdx; k-- {
			newWidths[toColLabel(k+1)] = s.ColWidths[toColLabel(k)]
		}
		newWidths[toColLabel(destIdx)] = s.ColWidths[toColLabel(fromIdx)]
	}
	s.ColWidths = newWidths

	// Audit entry
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "MOVE_COL",
		Details:   fmt.Sprintf("Moved column %s to right of column %s", fromColStr, targetColStr),
	})

	s.mu.Unlock()

	// Save after unlock
	globalSheetManager.SaveSheet(s)
	return true
}

// InsertColumnRight inserts a new empty column directly to the right of `targetColStr`,
// shifting subsequent columns (data and widths) right by one. Returns true if insertion occurred.
func (s *Sheet) InsertColumnRight(targetColStr, user string) bool {
	// Reuse the same helpers as MoveColumnRight
	toColIdx := func(label string) int {
		idx := 0
		for i := 0; i < len(label); i++ {
			idx = idx*26 + int(label[i]-'A'+1)
		}
		return idx
	}
	toColLabel := func(idx int) string {
		label := ""
		for idx > 0 {
			idx--
			b := byte(int('A') + (idx % 26))
			label = string([]byte{b}) + label
			idx /= 26
		}
		return label
	}

	targetIdx := toColIdx(targetColStr)
	if targetIdx == 0 {
		return false
	}
	insertIdx := targetIdx + 1

	s.mu.Lock()

	// Determine current max column index based on ColWidths and Data
	maxIdx := 0
	for col := range s.ColWidths {
		if idx := toColIdx(col); idx > maxIdx {
			maxIdx = idx
		}
	}
	for _, rowMap := range s.Data {
		for col := range rowMap {
			if idx := toColIdx(col); idx > maxIdx {
				maxIdx = idx
			}
		}
	}

	// Shift cells for all rows from right to left
	for idx := maxIdx; idx >= insertIdx; idx-- {
		fromLabel := toColLabel(idx)
		toLabel := toColLabel(idx + 1)
		for rowKey, rowMap := range s.Data {
			if cell, ok := rowMap[fromLabel]; ok {
				if s.Data[rowKey] == nil {
					s.Data[rowKey] = make(map[string]Cell)
				}
				rowMap[toLabel] = cell
				delete(rowMap, fromLabel)
			} else {
				delete(rowMap, toLabel)
			}
		}
	}

	// Ensure the new column exists as empty in all rows (optional but consistent)
	newLabel := toColLabel(insertIdx)
	for rowKey := range s.Data {
		if s.Data[rowKey] == nil {
			s.Data[rowKey] = make(map[string]Cell)
		}
		if _, ok := s.Data[rowKey][newLabel]; !ok {
			s.Data[rowKey][newLabel] = Cell{}
		}
	}

	// Shift ColWidths
	if s.ColWidths == nil {
		s.ColWidths = make(map[string]int)
	}
	maxWidthIdx := 0
	for col := range s.ColWidths {
		if idx := toColIdx(col); idx > maxWidthIdx {
			maxWidthIdx = idx
		}
	}
	for idx := maxWidthIdx; idx >= insertIdx; idx-- {
		fromLabel := toColLabel(idx)
		toLabel := toColLabel(idx + 1)
		if w, ok := s.ColWidths[fromLabel]; ok {
			delete(s.ColWidths, fromLabel)
			s.ColWidths[toLabel] = w
		} else {
			delete(s.ColWidths, toLabel)
		}
	}
	// New column width defaults to existing width of target column, if any
	if w, ok := s.ColWidths[targetColStr]; ok {
		s.ColWidths[newLabel] = w
	}

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "INSERT_COL",
		Details:   "Inserted column " + newLabel + " to the right of column " + targetColStr,
	})

	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
	return true
}
func itoa(i int) string {
	return fmt.Sprintf("%d", i)
}

func (sm *SheetManager) ListSheets() []*Sheet {
	sm.mu.RLock()
	defer sm.mu.RUnlock()

	list := make([]*Sheet, 0, len(sm.sheets))
	for _, sheet := range sm.sheets {
		list = append(list, sheet)
	}
	return list
}

func (sm *SheetManager) RenameSheet(id, newName, user string) bool {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	// Find sheet across projects
	var sheet *Sheet
	for _, s := range sm.sheets {
		if s != nil && s.ID == id {
			sheet = s
			break
		}
	}
	if sheet == nil {
		return false
	}

	sheet.mu.Lock()
	oldName := sheet.Name
	sheet.Name = newName
	sheet.mu.Unlock()

	// Project-level audit only
	globalProjectAuditManager.Append(sheet.ProjectName, user, "RENAME_SHEET", "Renamed sheet from '"+oldName+"' to '"+newName+"'")

	// Persist with existing key
	sm.saveSheetLocked(sheet)
	return true
}

// RenameSheetBy renames a sheet identified by id and project.
func (sm *SheetManager) RenameSheetBy(id, project, newName, user string) bool {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	var sheet *Sheet
	// Prefer composite key match
	if s, ok := sm.sheets[sheetKey(project, id)]; ok {
		sheet = s
	} else {
		for _, s := range sm.sheets {
			if s != nil && s.ID == id && s.ProjectName == project {
				sheet = s
				break
			}
		}
	}
	if sheet == nil {
		return false
	}

	sheet.mu.Lock()
	oldName := sheet.Name
	sheet.Name = newName
	sheet.mu.Unlock()

	// Project-level audit only
	globalProjectAuditManager.Append(project, user, "RENAME_SHEET", "Renamed sheet from '"+oldName+"' to '"+newName+"'")

	// Persist
	sm.saveSheetLocked(sheet)
	return true
}

func (sm *SheetManager) DeleteSheet(id string) bool {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	// Find entry by id
	var sheet *Sheet
	for _, s := range sm.sheets {
		if s != nil && s.ID == id {
			sheet = s
			break
		}
	}
	if sheet == nil {
		return false
	}

	// Delete using computed composite key
	delete(sm.sheets, sheetKey(sheet.ProjectName, id))

	// Remove the sheet file
	var filePath string
	if sheet.ProjectName != "" {
		filePath = filepath.Join(dataDir, sheet.ProjectName, id+".json")
	} else {
		filePath = getSheetFilePath(id)
	}
	if err := os.Remove(filePath); err != nil {
		log.Printf("Error deleting sheet file %s: %v", filePath, err)
	}

	return true
}

// DeleteSheetBy deletes a sheet with id and project from memory and disk.
func (sm *SheetManager) DeleteSheetBy(id, project string) bool {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	var sheet *Sheet
	if s, ok := sm.sheets[sheetKey(project, id)]; ok {
		sheet = s
	} else {
		for _, s := range sm.sheets {
			if s != nil && s.ID == id && s.ProjectName == project {
				sheet = s
				break
			}
		}
	}
	if sheet == nil {
		return false
	}

	delete(sm.sheets, sheetKey(project, id))

	// Remove the sheet file
	var filePath string
	if sheet.ProjectName != "" {
		filePath = filepath.Join(dataDir, sheet.ProjectName, id+".json")
	} else {
		filePath = getSheetFilePath(id)
	}
	if err := os.Remove(filePath); err != nil {
		log.Printf("Error deleting sheet file %s: %v", filePath, err)
	}

	return true
}

// DeleteSheetsByProject removes all sheets in a given project from memory and disk.
func (sm *SheetManager) DeleteSheetsByProject(projectName string) {
	sm.mu.Lock()
	// Collect ids to delete to avoid mutating map during iteration
	ids := make([]string, 0)
	for id, s := range sm.sheets {
		if s.ProjectName == projectName {
			ids = append(ids, id)
		}
	}
	sm.mu.Unlock()
	for _, id := range ids {
		sm.DeleteSheet(id)
	}
}

func (sm *SheetManager) SaveSheet(sheet *Sheet) {
	sm.mu.RLock()
	defer sm.mu.RUnlock()
	sm.saveSheetLocked(sheet)
}

func (sm *SheetManager) Save() {
	sm.mu.RLock()
	defer sm.mu.RUnlock()
	// Save all sheets
	for _, sheet := range sm.sheets {
		sm.saveSheetLocked(sheet)
	}
}

func (sm *SheetManager) Load() {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	// Check if DATA directory exists
	if _, err := os.Stat(dataDir); os.IsNotExist(err) {
		log.Println("DATA directory does not exist, starting fresh")
		return
	}

	loadedCount := 0

	// Load sheets from root (backward compatibility)
	rootFiles, err := filepath.Glob(filepath.Join(dataDir, "*.json"))
	if err == nil {
		for _, filePath := range rootFiles {
			base := filepath.Base(filePath)
			// Skip non-sheet files like chat.json, projects.json, users.json
			if base == "chat.json" || base == "projects.json" || base == "users.json" {
				continue
			}
			file, err := os.Open(filePath)
			if err != nil {
				log.Printf("Error opening sheet file %s: %v", filePath, err)
				continue
			}
			var sheet Sheet
			if err := json.NewDecoder(file).Decode(&sheet); err != nil {
				log.Printf("Error decoding sheet file %s: %v", filePath, err)
				file.Close()
				continue
			}
			file.Close()
			sm.sheets[sheetKey(sheet.ProjectName, sheet.ID)] = &sheet
			loadedCount++
		}
	}

	// Load sheets from project subdirectories
	entries, err := os.ReadDir(dataDir)
	if err != nil {
		log.Printf("Error reading DATA directory: %v", err)
	} else {
		for _, entry := range entries {
			if !entry.IsDir() {
				continue
			}
			project := entry.Name()
			files, err := filepath.Glob(filepath.Join(dataDir, project, "*.json"))
			if err != nil {
				log.Printf("Error listing files for project %s: %v", project, err)
				continue
			}
			for _, filePath := range files {
				file, err := os.Open(filePath)
				if err != nil {
					log.Printf("Error opening sheet file %s: %v", filePath, err)
					continue
				}
				var sheet Sheet
				if err := json.NewDecoder(file).Decode(&sheet); err != nil {
					log.Printf("Error decoding sheet file %s: %v", filePath, err)
					file.Close()
					continue
				}
				file.Close()
				// If project name missing in file, infer from folder
				if sheet.ProjectName == "" {
					sheet.ProjectName = project
				}
				sm.sheets[sheetKey(sheet.ProjectName, sheet.ID)] = &sheet
				loadedCount++
			}
		}
	}

	log.Printf("Loaded %d sheets from disk", loadedCount)
}

// DuplicateProject duplicates all sheets in source project into a new project name.
func (sm *SheetManager) DuplicateProject(sourceProject, newProject string) error {
	if sourceProject == "" || newProject == "" {
		return fmt.Errorf("source and new project names required")
	}
	// Ensure destination directory
	destDir := filepath.Join(dataDir, newProject)
	if err := os.MkdirAll(destDir, 0755); err != nil {
		return fmt.Errorf("failed to create dest project dir: %w", err)
	}
	// Gather sheets to duplicate
	sm.mu.Lock()
	defer sm.mu.Unlock()
	for _, s := range sm.sheets {
		if s == nil || s.ProjectName != sourceProject {
			continue
		}
		// Clone maintaining same ID and owner/permissions
		clone := &Sheet{
			ID:          s.ID,
			Name:        s.Name,
			Owner:       s.Owner,
			ProjectName: newProject,
			Data:        make(map[string]map[string]Cell),
			ColWidths:   make(map[string]int),
			RowHeights:  make(map[string]int),
			Permissions: s.Permissions,
			AuditLog:    append([]AuditEntry{}, s.AuditLog...),
		}
		// Deep copy state
		s.mu.RLock()
		for r, cols := range s.Data {
			clone.Data[r] = make(map[string]Cell, len(cols))
			for c, cell := range cols {
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
		// Register and persist
		sm.sheets[sheetKey(newProject, clone.ID)] = clone
		sm.saveSheetLocked(clone)
	}
	return nil
}
