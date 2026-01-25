package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sync"
	"time"
)

const dataDir = "../DATA"

func firstNChar(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n] + "..."
}

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
	Timestamp      time.Time `json:"timestamp"`
	User           string    `json:"user"`
	Action         string    `json:"action"` // e.g., "EDIT_CELL", "CREATE_SHEET"
	Details        string    `json:"details,omitempty"`
	Row1           int       `json:"row,omitempty"`
	Col1           string    `json:"col,omitempty"`
	Row2           int       `json:"row,omitempty"`
	Col2           string    `json:"col,omitempty"`
	OldValue       string    `json:"old_value,omitempty"`       // Added for tracking previous value
	NewValue       string    `json:"new_value,omitempty"`       // Added for tracking new value
	ChangeReversed bool      `json:"change_reversed,omitempty"` // Indicates if the change was reversed by owner later default value is false
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

	// Async save queue (debounced per sheet)
	pending      map[string]*pendingSave // key -> pending info
	saveInterval time.Duration           // debounce duration
	started      bool
	stopCh       chan struct{}
}

type pendingSave struct {
	sheet        *Sheet
	lastModified time.Time
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
	sheets:  make(map[string]*Sheet),
	pending: make(map[string]*pendingSave),
	stopCh:  make(chan struct{}),
	started: false,
	// saveInterval will be set in initAsyncSaver
}

// Initialize async saver once during startup
func init() {
	globalSheetManager.initAsyncSaver()
}

func (sm *SheetManager) initAsyncSaver() {
	sm.mu.Lock()
	defer sm.mu.Unlock()
	if sm.started {
		return
	}
	if sm.saveInterval == 0 {
		sm.saveInterval = 1 * time.Second // default debounce window
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
		case <-sm.stopCh:
			return
		}
	}
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

	// Initial Audit (details left empty for persistence)
	sheet.AuditLog = append(sheet.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      owner,
		Action:    "CREATE_SHEET",
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
			Timestamp:      time.Now(),
			User:           user,
			Action:         "EDIT_CELL",
			Row1:           atoiSafe(row),
			Col1:           col,
			OldValue:       currentVal.Value,
			NewValue:       value,
			ChangeReversed: false,
		})
	} else {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp:      time.Now(),
			User:           user,
			Action:         "EDIT_CELL",
			Row1:           atoiSafe(row),
			Col1:           col,
			OldValue:       "",
			NewValue:       value,
			ChangeReversed: false,
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
			Timestamp:      time.Now(),
			User:           user,
			Action:         "STYLE_CELL",
			Row1:           atoiSafe(row),
			Col1:           col,
			ChangeReversed: false,
		})
	} else {
		s.AuditLog = append(s.AuditLog, AuditEntry{
			Timestamp:      time.Now(),
			User:           user,
			Action:         "STYLE_CELL",
			Row1:           atoiSafe(row),
			Col1:           col,
			ChangeReversed: false,
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
		Timestamp:      time.Now(),
		User:           user,
		Action:         "LOCK_CELL",
		Row1:           atoiSafe(row),
		Col1:           col,
		ChangeReversed: false,
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
		Timestamp:      time.Now(),
		User:           user,
		Action:         "UNLOCK_CELL",
		Row1:           atoiSafe(row),
		Col1:           col,
		ChangeReversed: false,
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

	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
}

func (s *Sheet) SetRowHeight(row string, height int, user string) {
	s.mu.Lock()
	if s.RowHeights == nil {
		s.RowHeights = make(map[string]int)
	}
	s.RowHeights[row] = height

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
	go globalSheetManager.SaveSheet(s)
	// Log only in project audit
	globalProjectAuditManager.Append(s.ProjectName, performedBy, "UPDATE_SHEET_PERMISSIONS", fmt.Sprintf("For Sheet %s Editors: %v", s.ID, editors))
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
	go globalSheetManager.SaveSheet(s)
	// Log only in project audit
	globalProjectAuditManager.Append(s.ProjectName, performedBy, "TRANSFER_SHEET_OWNERSHIP", fmt.Sprintf("For Sheet %s Owner changed from %s to %s", s.ID, old, newOwner))
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
	// Adjust audit log row references for rows at or below the inserted position
	s.adjustAuditRowsOnInsert(insertRow)

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "INSERT_ROW",
		Row1:      insertRow,
	})

	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
	return true
}

// DeleteRowAt removes the row at rowStr and shifts subsequent rows up by one
func (s *Sheet) DeleteRowAt(rowStr, user string) bool {
	var row int
	if _, err := fmt.Sscanf(rowStr, "%d", &row); err != nil || row <= 0 {
		return false
	}
	s.mu.Lock()
	// Determine max row
	maxRow := 0
	for rowKey := range s.Data {
		var r int
		if _, err := fmt.Sscanf(rowKey, "%d", &r); err == nil {
			if r > maxRow {
				maxRow = r
			}
		}
	}
	// Remove the target row
	delete(s.Data, rowStr)
	// Shift rows [row+1..maxRow] up by 1
	for r := row + 1; r <= maxRow; r++ {
		fromKey := itoa(r)
		toKey := itoa(r - 1)
		if rowData, ok := s.Data[fromKey]; ok {
			delete(s.Data, fromKey)
			s.Data[toKey] = rowData
		} else {
			delete(s.Data, toKey)
		}
	}
	// RowHeights shift
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
	delete(s.RowHeights, rowStr)
	for r := row + 1; r <= maxHeightRow; r++ {
		fromKey := itoa(r)
		toKey := itoa(r - 1)
		if h, ok := s.RowHeights[fromKey]; ok {
			delete(s.RowHeights, fromKey)
			s.RowHeights[toKey] = h
		} else {
			delete(s.RowHeights, toKey)
		}
	}
	// Adjust audit logs for deletion
	s.adjustAuditRowsOnDelete(row)
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "DELETE_ROW",
		Row1:      row,
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
		Row1:      fromRow,
		Row2:      destIndex,
	})

	// Adjust audit log rows according to move mapping
	s.adjustAuditRowsOnMove(fromRow, destIndex)

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
	for _, rowMap := range s.Data {
		if cell, ok := rowMap[fromColStr]; ok {
			if cell.Locked {
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
		Col1:      fromColStr,
		Col2:      toColLabel(destIdx),
	})

	// Adjust audit log columns according to move mapping
	s.adjustAuditColsOnMove(fromIdx, destIdx)

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
	// Adjust audit log column references for columns at or beyond the inserted position
	s.adjustAuditColsOnInsert(insertIdx)

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp:      time.Now(),
		User:           user,
		Action:         "INSERT_COL",
		Col1:           newLabel,
		ChangeReversed: false,
	})

	s.mu.Unlock()

	globalSheetManager.SaveSheet(s)
	return true
}

// DeleteColumnAt removes a column by label and shifts subsequent columns left by one
func (s *Sheet) DeleteColumnAt(colStr, user string) bool {
	insertIdx := colLabelToIndex(colStr)
	if insertIdx <= 0 {
		return false
	}
	s.mu.Lock()
	// Determine max column index
	maxIdx := 0
	for _, rowMap := range s.Data {
		for col := range rowMap {
			if idx := colLabelToIndex(col); idx > maxIdx {
				maxIdx = idx
			}
		}
	}
	// Remove target column
	for _, rowMap := range s.Data {
		delete(rowMap, colStr)
	}
	// Shift [insertIdx+1..maxIdx] left by 1
	for idx := insertIdx + 1; idx <= maxIdx; idx++ {
		fromLabel := indexToColLabel(idx)
		toLabel := indexToColLabel(idx - 1)
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
	// Shift ColWidths
	if s.ColWidths == nil {
		s.ColWidths = make(map[string]int)
	}
	maxWidthIdx := 0
	for col := range s.ColWidths {
		if idx := colLabelToIndex(col); idx > maxWidthIdx {
			maxWidthIdx = idx
		}
	}
	delete(s.ColWidths, colStr)
	for idx := insertIdx + 1; idx <= maxWidthIdx; idx++ {
		fromLabel := indexToColLabel(idx)
		toLabel := indexToColLabel(idx - 1)
		if w, ok := s.ColWidths[fromLabel]; ok {
			delete(s.ColWidths, fromLabel)
			s.ColWidths[toLabel] = w
		} else {
			delete(s.ColWidths, toLabel)
		}
	}
	// Adjust audit logs for deletion
	s.adjustAuditColsOnDelete(insertIdx)
	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp:      time.Now(),
		User:           user,
		Action:         "DELETE_COL",
		Col1:           colStr,
		ChangeReversed: false,
	})

	s.mu.Unlock()
	globalSheetManager.SaveSheet(s)
	return true
}
func itoa(i int) string {
	return fmt.Sprintf("%d", i)
}

func atoiSafe(s string) int {
	var v int
	if _, err := fmt.Sscanf(s, "%d", &v); err != nil {
		return 0
	}
	return v
}

// Column helpers
func colLabelToIndex(label string) int {
	if label == "" {
		return 0
	}
	idx := 0
	for i := 0; i < len(label); i++ {
		ch := label[i]
		if ch < 'A' || ch > 'Z' {
			return 0
		}
		idx = idx*26 + int(ch-'A'+1)
	}
	return idx
}

func indexToColLabel(idx int) string {
	if idx <= 0 {
		return ""
	}
	label := ""
	for idx > 0 {
		idx--
		b := byte(int('A') + (idx % 26))
		label = string([]byte{b}) + label
		idx /= 26
	}
	return label
}

var (
	reCell = regexp.MustCompile(`(?i)cell\s+(\d+),([A-Z]+)`) // cell 12,C
	reRow  = regexp.MustCompile(`(?i)row\s+(\d+)`)           // row 12
	reCol  = regexp.MustCompile(`(?i)column\s+([A-Z]+)`)     // column C
)

// ensureEntryCoords tries to ensure Row/Col fields are populated for an entry by parsing Details if needed
func ensureEntryCoords(e *AuditEntry) {
	if e.Row1 == 0 || e.Col1 == "" {
		if m := reCell.FindStringSubmatch(e.Details); len(m) == 3 {
			var r int
			_, _ = fmt.Sscanf(m[1], "%d", &r)
			e.Row1 = r
			e.Col1 = m[2]
			return
		}
	}
	if e.Row1 == 0 {
		if m := reRow.FindStringSubmatch(e.Details); len(m) == 2 {
			var r int
			_, _ = fmt.Sscanf(m[1], "%d", &r)
			e.Row1 = r
		}
	}
	if e.Col1 == "" {
		if m := reCol.FindStringSubmatch(e.Details); len(m) == 2 {
			e.Col1 = m[1]
		}
	}
}

// replaceDetailCoords updates the coordinates present in Details if patterns are found
func replaceDetailCoords(details string, newRow int, newCol string) string {
	updated := details
	// Replace only the first occurrence of row and column to reflect the entry's coordinates
	if newRow > 0 {
		if loc := reRow.FindStringIndex(updated); loc != nil {
			// Replace entire match with "row <newRow>"
			updated = updated[:loc[0]] + fmt.Sprintf("row %d", newRow) + updated[loc[1]:]
		}
	}
	if newCol != "" {
		if loc := reCol.FindStringIndex(updated); loc != nil {
			updated = updated[:loc[0]] + fmt.Sprintf("column %s", newCol) + updated[loc[1]:]
		}
	}
	if newRow > 0 || newCol != "" {
		if reCell.MatchString(updated) {
			updated = reCell.ReplaceAllStringFunc(updated, func(s string) string {
				// Ignore s; replace using provided newRow/newCol when available
				// If one of them is missing, preserve original via captured groups
				m := reCell.FindStringSubmatch(s)
				rowStr, colStr := m[1], m[2]
				if newRow > 0 {
					rowStr = itoa(newRow)
				}
				if newCol != "" {
					colStr = newCol
				}
				return fmt.Sprintf("cell %s,%s", rowStr, colStr)
			})
		}
	}
	return updated
}

// computeAuditDetails constructs a user-friendly details string for an audit entry
// without persisting it. Uses structured fields from the entry and, when needed,
// sheet context (e.g., sheet name).
func computeAuditDetails(s *Sheet, e AuditEntry) string {
	switch e.Action {
	case "CREATE_SHEET":
		if s != nil {
			return "Created sheet " + s.Name
		}
		return "Created sheet"
	case "EDIT_CELL":
		r := e.Row1
		c := e.Col1
		if e.OldValue == "" {
			return fmt.Sprintf("Set cell %d,%s to %s", r, c, firstNChar(e.NewValue, 10))
		}
		return fmt.Sprintf("Changed cell %d,%s from %s to %s", r, c, firstNChar(e.OldValue, 10), firstNChar(e.NewValue, 10))
	case "STYLE_CELL":
		return fmt.Sprintf("Updated style for cell %d,%s", e.Row1, e.Col1)
	case "LOCK_CELL":
		return fmt.Sprintf("Locked cell %d,%s", e.Row1, e.Col1)
	case "UNLOCK_CELL":
		return fmt.Sprintf("Unlocked cell %d,%s", e.Row1, e.Col1)
	case "INSERT_ROW":
		return fmt.Sprintf("Inserted row %d", e.Row1)
	case "DELETE_ROW":
		return fmt.Sprintf("Deleted row %d", e.Row1)
	case "MOVE_ROW":
		if e.Row2 > 0 {
			return fmt.Sprintf("Moved row %d to row %d", e.Row1, e.Row2)
		}
		return fmt.Sprintf("Moved row %d", e.Row1)
	case "INSERT_COL":
		return fmt.Sprintf("Inserted column %s", e.Col1)
	case "DELETE_COL":
		return fmt.Sprintf("Deleted column %s", e.Col1)
	case "MOVE_COL":
		if e.Col2 != "" {
			return fmt.Sprintf("Moved column %s to right of column %s", e.Col1, e.Col2)
		}
		return fmt.Sprintf("Moved column %s", e.Col1)
	case "UPDATE_PERMISSIONS":
		return "Updated permissions"
	case "TRANSFER_OWNERSHIP":
		return "Transferred ownership"
	default:
		return e.Action
	}
}

// SnapshotForClient builds a copy of the sheet with audit Details filled for response
// without mutating or leaking internal state. This snapshot is safe to marshal/send.
func (s *Sheet) SnapshotForClient() *Sheet {
	s.mu.RLock()
	defer s.mu.RUnlock()

	// Deep copy data
	dataCopy := make(map[string]map[string]Cell, len(s.Data))
	for r, cols := range s.Data {
		inner := make(map[string]Cell, len(cols))
		for c, cell := range cols {
			inner[c] = cell
		}
		dataCopy[r] = inner
	}
	colWidthsCopy := make(map[string]int, len(s.ColWidths))
	for k, v := range s.ColWidths {
		colWidthsCopy[k] = v
	}
	rowHeightsCopy := make(map[string]int, len(s.RowHeights))
	for k, v := range s.RowHeights {
		rowHeightsCopy[k] = v
	}
	auditCopy := make([]AuditEntry, 0, len(s.AuditLog))
	for _, e := range s.AuditLog {
		e2 := e
		e2.Details = computeAuditDetails(s, e)
		auditCopy = append(auditCopy, e2)
	}

	snap := &Sheet{
		ID:          s.ID,
		Name:        s.Name,
		Owner:       s.Owner,
		ProjectName: s.ProjectName,
		Data:        dataCopy,
		AuditLog:    auditCopy,
		Permissions: Permissions{Editors: append([]string(nil), s.Permissions.Editors...)},
		ColWidths:   colWidthsCopy,
		RowHeights:  rowHeightsCopy,
	}
	return snap
}

// adjustAuditRowsOnInsert increments row references for entries at or below insertRow
func (s *Sheet) adjustAuditRowsOnInsert(insertRow int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		oldRow := e.Row1
		ensureEntryCoords(e)
		if e.Row1 >= insertRow && e.Row1 > 0 {
			e.Row1 = e.Row1 + 1
		}
		_ = oldRow // details left empty; no string rewrite
	}
}

// adjustAuditRowsOnMove adjusts row references for a row move from fromRow to destIndex
func (s *Sheet) adjustAuditRowsOnMove(fromRow, destIndex int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		ensureEntryCoords(e)
		if e.Row1 == 0 {
			continue
		}
		old := e.Row1
		if fromRow < destIndex {
			if e.Row1 == fromRow {
				e.Row1 = destIndex
			} else if e.Row1 > fromRow && e.Row1 <= destIndex {
				e.Row1 = e.Row1 - 1
			}
		} else if fromRow > destIndex {
			if e.Row1 == fromRow {
				e.Row1 = destIndex
			} else if e.Row1 >= destIndex && e.Row1 < fromRow {
				e.Row1 = e.Row1 + 1
			}
		}
		_ = old // details left empty; no string rewrite
	}
}

// adjustAuditColsOnInsert increments column references for entries at or beyond insertIdx
func (s *Sheet) adjustAuditColsOnInsert(insertIdx int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		oldCol := e.Col1
		ensureEntryCoords(e)
		idx := colLabelToIndex(e.Col1)
		if idx >= insertIdx && idx > 0 {
			idx = idx + 1
			e.Col1 = indexToColLabel(idx)
		}
		_ = oldCol // details left empty; no string rewrite
	}
}

// adjustAuditColsOnMove adjusts column references for a column move from fromIdx to destIdx
func (s *Sheet) adjustAuditColsOnMove(fromIdx, destIdx int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		ensureEntryCoords(e)
		idx := colLabelToIndex(e.Col1)
		if idx == 0 {
			continue
		}
		oldIdx := idx
		if fromIdx < destIdx {
			if idx == fromIdx {
				idx = destIdx
			} else if idx > fromIdx && idx <= destIdx {
				idx = idx - 1
			}
		} else if fromIdx > destIdx {
			if idx == fromIdx {
				idx = destIdx
			} else if idx >= destIdx && idx < fromIdx {
				idx = idx + 1
			}
		}
		if idx != oldIdx {
			e.Col1 = indexToColLabel(idx)
		}
	}
}

// adjustAuditRowsOnDelete decrements row references for entries strictly above deleted row
func (s *Sheet) adjustAuditRowsOnDelete(deleteRow int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		oldRow := e.Row1
		ensureEntryCoords(e)
		if e.Row1 > deleteRow {
			e.Row1 = e.Row1 - 1
		}
		_ = oldRow // details left empty; no string rewrite
	}
}

// adjustAuditColsOnDelete decrements column references for entries strictly right of deleted column
func (s *Sheet) adjustAuditColsOnDelete(deleteIdx int) {
	for i := range s.AuditLog {
		e := &s.AuditLog[i]
		oldCol := e.Col1
		ensureEntryCoords(e)
		idx := colLabelToIndex(e.Col1)
		if idx > deleteIdx {
			idx = idx - 1
			e.Col1 = indexToColLabel(idx)
		}
		_ = oldCol // details left empty; no string rewrite
	}
}

// (Removed) MoveColumnToIndex: undo now uses MOVE_COL with computed target column.

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
	// Schedule a debounced save instead of writing immediately
	if sheet == nil {
		return
	}
	// Build key from sheet fields safely
	sheet.mu.RLock()
	proj := sheet.ProjectName
	id := sheet.ID
	sheet.mu.RUnlock()

	key := sheetKey(proj, id)
	now := time.Now()
	sm.mu.Lock()
	if ps, ok := sm.pending[key]; ok {
		ps.lastModified = now
		// keep existing sheet pointer; it always refers to same instance
	} else {
		sm.pending[key] = &pendingSave{sheet: sheet, lastModified: now}
	}
	sm.mu.Unlock()
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
	/*
		rootFiles, err := filepath.Glob(filepath.Join(dataDir, "*.json"))
		if err == nil {
			for _, filePath := range rootFiles {
				base := filepath.Base(filePath)
				// Skip non-sheet files like chat.json, projects.json, users.json
				if base == "chat.json" || base == "projects.json" || base == "users.json" || base == "project_audit.log" {
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
	*/
	// Load sheets from project subdirectories (recursive)
	entries, err := os.ReadDir(dataDir)
	if err != nil {
		log.Printf("Error reading DATA directory: %v", err)
	} else {
		for _, entry := range entries {
			if !entry.IsDir() {
				continue
			}
			topProject := entry.Name()
			baseDir := filepath.Join(dataDir, topProject)
			// Walk recursively and read any *.json sheet file
			filepath.WalkDir(baseDir, func(path string, d os.DirEntry, err error) error {
				if err != nil {
					return nil
				}
				if d.IsDir() {
					return nil
				}
				if filepath.Ext(path) != ".json" {
					return nil
				}
				// Skip non-sheet meta files
				base := filepath.Base(path)
				if base == "chat.json" || base == "projects.json" || base == "users.json" || base == "project_audit.log" {
					return nil
				}
				file, err := os.Open(path)
				if err != nil {
					log.Printf("Error opening sheet file %s: %v", path, err)
					return nil
				}
				var sheet Sheet
				if err := json.NewDecoder(file).Decode(&sheet); err != nil {
					log.Printf("Error decoding sheet file %s: %v", path, err)
					file.Close()
					return nil
				}
				file.Close()
				// If project name missing in file, infer relative project path from DATA dir
				if sheet.ProjectName == "" {
					rel, relErr := filepath.Rel(dataDir, filepath.Dir(path))
					if relErr == nil {
						sheet.ProjectName = rel
					} else {
						sheet.ProjectName = topProject
					}
				}
				sm.sheets[sheetKey(sheet.ProjectName, sheet.ID)] = &sheet
				loadedCount++
				return nil
			})
		}
	}

	log.Printf("Loaded %d sheets from disk", loadedCount)
}

// DuplicateProject duplicates all sheets in source project into a new project name.
// The newOwner will be set as the owner for all duplicated sheets (and ensured in editors).
func (sm *SheetManager) DuplicateProject(sourceProject, newProject, newOwner string) error {
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
		// Ensure new owner is present in editors
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
			ID:          s.ID,
			Name:        s.Name,
			Owner:       newOwner,
			ProjectName: newProject,
			Data:        make(map[string]map[string]Cell),
			ColWidths:   make(map[string]int),
			RowHeights:  make(map[string]int),
			Permissions: perms,
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
