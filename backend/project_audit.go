package main

import (
	"encoding/json"
	"log"
	"os"
	"path/filepath"
	"sync"
	"time"
)

// Project-level audit entry
type ProjectAuditEntry struct {
	Timestamp time.Time `json:"timestamp"`
	Project   string    `json:"project"`
	User      string    `json:"user"`
	Action    string    `json:"action"`
	Details   string    `json:"details"`
}

// Manages per-project audit logs persisted to DATA/project_audit.json
type ProjectAuditManager struct {
	mu   sync.RWMutex
	logs map[string][]ProjectAuditEntry // project -> entries
}

var globalProjectAuditManager = &ProjectAuditManager{
	logs: make(map[string][]ProjectAuditEntry),
}

func (pm *ProjectAuditManager) filePath() string {
	return filepath.Join(dataDir, "project_audit.json")
}

func (pm *ProjectAuditManager) Load() {
	pm.mu.Lock()
	defer pm.mu.Unlock()
	// Ensure data dir exists
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		log.Printf("project audit: ensure data dir: %v", err)
	}
	absPath, _ := filepath.Abs(pm.filePath())
	data, err := os.ReadFile(pm.filePath())
	if err != nil {
		if os.IsNotExist(err) {
			pm.logs = make(map[string][]ProjectAuditEntry)
			return
		}
		log.Printf("project audit: read: %v", err)
		return
	}
	// Integrity check
	CheckAndRecord(absPath, data)
	var m map[string][]ProjectAuditEntry
	if err := json.Unmarshal(data, &m); err != nil {
		log.Printf("project audit: decode: %v", err)
		globalIntegrity.Record(absPath, false, false, "json decode error: "+err.Error())
		return
	}
	pm.logs = m
}

func (pm *ProjectAuditManager) Save() {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	// Ensure data dir exists
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		log.Printf("project audit: ensure data dir: %v", err)
		return
	}
	data, err := json.MarshalIndent(pm.logs, "", "  ")
	if err != nil {
		log.Printf("project audit: encode: %v", err)
		return
	}
	data = append(data, '\n')
	absPath, _ := filepath.Abs(pm.filePath())
	if err := WriteFileWithChecksum(absPath, data); err != nil {
		log.Printf("project audit: save: %v", err)
		return
	}
	globalIntegrity.Record(absPath, true, false, "")
}

func (pm *ProjectAuditManager) Append(project, user, action, details string) {
	if project == "" {
		return
	}
	entry := ProjectAuditEntry{
		Timestamp: time.Now(),
		Project:   project,
		User:      user,
		Action:    action,
		Details:   details,
	}
	pm.mu.Lock()
	pm.logs[project] = append(pm.logs[project], entry)
	pm.mu.Unlock()
	pm.Save()
}

func (pm *ProjectAuditManager) List(project string) []ProjectAuditEntry {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	return append([]ProjectAuditEntry{}, pm.logs[project]...)
}
