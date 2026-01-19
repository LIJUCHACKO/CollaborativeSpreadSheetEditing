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
	Owner string `json:"owner"`
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
	f, err := os.Open(pm.filePath())
	if err != nil {
		if os.IsNotExist(err) {
			pm.data = make(map[string]ProjectMeta)
			return
		}
		log.Printf("project meta: open: %v", err)
		return
	}
	defer f.Close()
	dec := json.NewDecoder(f)
	var m map[string]ProjectMeta
	if err := dec.Decode(&m); err != nil {
		log.Printf("project meta: decode: %v", err)
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
	f, err := os.Create(pm.filePath())
	if err != nil {
		log.Printf("project meta: create: %v", err)
		return
	}
	defer f.Close()
	enc := json.NewEncoder(f)
	enc.SetIndent("", "  ")
	if err := enc.Encode(pm.data); err != nil {
		log.Printf("project meta: encode: %v", err)
	}
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
	pm.data[project] = ProjectMeta{Owner: owner}
	pm.mu.Unlock()
	pm.Save()
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
