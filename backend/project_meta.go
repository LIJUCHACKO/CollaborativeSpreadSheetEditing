package main

import (
	"encoding/json"
	"log"
	"os"
	"path/filepath"
	"sync"
)

// ProjectMetaManager persists simple metadata for projects, currently just the owner.
// Stored at DATA/projects.json as { "projectName": { "owner": "username" }, ... }

type ProjectMeta struct {
	Owner  string   `json:"owner"`
	Admins []string `json:"admins,omitempty"` // additional project admins (besides the owner)
}

type ProjectMetaManager struct {
	mu   sync.RWMutex
	data map[string]ProjectMeta // projectName -> meta
}

var globalProjectMeta = &ProjectMetaManager{data: make(map[string]ProjectMeta)}

func (pm *ProjectMetaManager) filePath() string {
	return filepath.Join(dataDir, "projects.json")
}

func (pm *ProjectMetaManager) Load() {
	pm.mu.Lock()
	defer pm.mu.Unlock()
	// Ensure data dir exists
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		log.Printf("project meta: ensure data dir: %v", err)
	}
	absPath, _ := filepath.Abs(pm.filePath())
	data, err := os.ReadFile(pm.filePath())
	if err != nil {
		if os.IsNotExist(err) {
			pm.data = make(map[string]ProjectMeta)
			return
		}
		log.Printf("project meta: read: %v", err)
		return
	}
	// Integrity check
	CheckAndRecord(absPath, data)
	var m map[string]ProjectMeta
	if err := json.Unmarshal(data, &m); err != nil {
		log.Printf("project meta: decode: %v", err)
		globalIntegrity.Record(absPath, false, false, "json decode error: "+err.Error())
		return
	}
	pm.data = m
}

func (pm *ProjectMetaManager) Save() {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	if err := os.MkdirAll(dataDir, 0755); err != nil {
		log.Printf("project meta: ensure data dir: %v", err)
		return
	}
	data, err := json.MarshalIndent(pm.data, "", "  ")
	if err != nil {
		log.Printf("project meta: encode: %v", err)
		return
	}
	data = append(data, '\n')
	absPath, _ := filepath.Abs(pm.filePath())
	if err := WriteFileWithChecksum(absPath, data); err != nil {
		log.Printf("project meta: save: %v", err)
		return
	}
	globalIntegrity.Record(absPath, true, false, "")
}

func (pm *ProjectMetaManager) GetOwner(project string) string {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	return pm.data[project].Owner
}

func (pm *ProjectMetaManager) SetOwner(project, owner string) {
	if project == "" || owner == "" {
		return
	}
	pm.mu.Lock()
	meta := pm.data[project]
	meta.Owner = owner
	pm.data[project] = meta
	pm.mu.Unlock()
	pm.Save()
}

// GetAdmins returns the list of additional admins for a project.
func (pm *ProjectMetaManager) GetAdmins(project string) []string {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	return pm.data[project].Admins
}

// SetAdmins replaces the admins list for a project.
func (pm *ProjectMetaManager) SetAdmins(project string, admins []string) {
	pm.mu.Lock()
	meta := pm.data[project]
	meta.Admins = admins
	pm.data[project] = meta
	pm.mu.Unlock()
	pm.Save()
}

// AddAdmin adds an admin to a project (no duplicates).
func (pm *ProjectMetaManager) AddAdmin(project, admin string) {
	if project == "" || admin == "" {
		return
	}
	pm.mu.Lock()
	meta := pm.data[project]
	for _, a := range meta.Admins {
		if a == admin {
			pm.mu.Unlock()
			return
		}
	}
	meta.Admins = append(meta.Admins, admin)
	pm.data[project] = meta
	pm.mu.Unlock()
	pm.Save()
}

// RemoveAdmin removes an admin from a project.
func (pm *ProjectMetaManager) RemoveAdmin(project, admin string) {
	pm.mu.Lock()
	meta := pm.data[project]
	newAdmins := make([]string, 0, len(meta.Admins))
	for _, a := range meta.Admins {
		if a != admin {
			newAdmins = append(newAdmins, a)
		}
	}
	meta.Admins = newAdmins
	pm.data[project] = meta
	pm.mu.Unlock()
	pm.Save()
}

// IsProjectAdmin returns true if the user is the project owner or one of the admins.
func (pm *ProjectMetaManager) IsProjectAdmin(project, user string) bool {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	meta := pm.data[project]
	if meta.Owner == user {
		return true
	}
	for _, a := range meta.Admins {
		if a == user {
			return true
		}
	}
	return false
}

func (pm *ProjectMetaManager) Delete(project string) {
	pm.mu.Lock()
	delete(pm.data, project)
	pm.mu.Unlock()
	pm.Save()
}

func (pm *ProjectMetaManager) Rename(oldName, newName string) {
	pm.mu.Lock()
	meta, ok := pm.data[oldName]
	if ok {
		delete(pm.data, oldName)
		pm.data[newName] = meta
	}
	pm.mu.Unlock()
	pm.Save()
}
