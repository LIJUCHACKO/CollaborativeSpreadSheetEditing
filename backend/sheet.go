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
	Value string `json:"value"`
	User  string `json:"user,omitempty"` // Last edited by
}

type AuditEntry struct {
	Timestamp time.Time `json:"timestamp"`
	User      string    `json:"user"`
	Action    string    `json:"action"` // e.g., "EDIT_CELL", "CREATE_SHEET"
	Details   string    `json:"details"`
}

type Permissions struct {
	Editors []string `json:"editors"`
	Viewers []string `json:"viewers"` // In this simple model, assume public read if empty, or specific list
}

type Sheet struct {
	ID          string                     `json:"id"`
	Name        string                     `json:"name"`
	Owner       string                     `json:"owner"`
	Data        map[string]map[string]Cell `json:"data"` // Row -> Col -> Cell
	AuditLog    []AuditEntry               `json:"audit_log"`
	Permissions Permissions                `json:"permissions"`
	ColWidths   map[string]int             `json:"col_widths,omitempty"`
	RowHeights  map[string]int             `json:"row_heights,omitempty"`
	mu          sync.RWMutex
}

type SheetManager struct {
	sheets map[string]*Sheet
	mu     sync.RWMutex
}

// Helper to save a single sheet without locking the manager (caller must hold lock)
func (sm *SheetManager) saveSheetLocked(sheet *Sheet) {
	if err := ensureDataDir(); err != nil {
		log.Printf("Error creating data directory: %v", err)
		return
	}

	filePath := getSheetFilePath(sheet.ID)
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

func (sm *SheetManager) CreateSheet(name, owner string) *Sheet {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	id := generateID() // Need to implement this or use a simple counter
	sheet := &Sheet{
		ID:         id,
		Name:       name,
		Owner:      owner,
		Data:       make(map[string]map[string]Cell),
		ColWidths:  make(map[string]int),
		RowHeights: make(map[string]int),
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

	sm.sheets[id] = sheet
	sm.saveSheetLocked(sheet) // Persist individual sheet
	return sheet
}

func (sm *SheetManager) GetSheet(id string) *Sheet {
	sm.mu.RLock()
	defer sm.mu.RUnlock()
	return sm.sheets[id]
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
	s.Data[row][col] = Cell{Value: value, User: user}

	s.AuditLog = append(s.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "EDIT_CELL",
		Details:   "Set cell " + row + "," + col + " to " + value,
	})

	s.mu.Unlock() // Unlock BEFORE saving to avoid deadlock (Save -> MarshalJSON -> tries RLock)

	// Persist changes
	// Optimally we shouldn't save on every cell edit for performance, but for this task it ensures safety.
	globalSheetManager.SaveSheet(s)
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

	sheet, exists := sm.sheets[id]
	if !exists {
		return false
	}

	sheet.mu.Lock()
	oldName := sheet.Name
	sheet.Name = newName
	sheet.AuditLog = append(sheet.AuditLog, AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    "RENAME_SHEET",
		Details:   "Renamed sheet from '" + oldName + "' to '" + newName + "'",
	})
	sheet.mu.Unlock()

	sm.saveSheetLocked(sheet)
	return true
}

func (sm *SheetManager) DeleteSheet(id string) bool {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	if _, exists := sm.sheets[id]; !exists {
		return false
	}

	delete(sm.sheets, id)

	// Remove the sheet file
	filePath := getSheetFilePath(id)
	if err := os.Remove(filePath); err != nil {
		log.Printf("Error deleting sheet file %s: %v", filePath, err)
	}

	return true
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

	// Read all .json files from DATA directory
	files, err := filepath.Glob(filepath.Join(dataDir, "*.json"))
	if err != nil {
		log.Printf("Error reading DATA directory: %v", err)
		return
	}

	loadedCount := 0
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

		sm.sheets[sheet.ID] = &sheet
		loadedCount++
	}

	log.Printf("Loaded %d sheets from disk", loadedCount)
}
