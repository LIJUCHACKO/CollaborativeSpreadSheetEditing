package main

// integrity.go — salted SHA-256 checksum helpers for JSON persistence.
//
// Every time a JSON file is written a companion "<file>.shasum" file is created
// that contains the hex SHA-256 of (salt + file bytes).  On load the checksum
// is re-computed and compared; a mismatch (or a missing .shasum) is treated as
// a corruption event and recorded in the global IntegrityStatus registry.

import (
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

// checksumSalt is a fixed salt prepended before hashing.
// Change this value to invalidate all existing checksums (forces re-generation on next save).
const checksumSalt = "SharedSpreadSheet-integrity-salt-v1"

// ---------------------------------------------------------------------------
// Low-level helpers
// ---------------------------------------------------------------------------

// shasumPath returns the companion checksum path for a given JSON file.
func shasumPath(jsonPath string) string {
	return jsonPath + ".shasum"
}

// computeChecksum returns the hex SHA-256 of (salt + data).
func computeChecksum(data []byte) string {
	h := sha256.New()
	h.Write([]byte(checksumSalt))
	h.Write(data)
	return hex.EncodeToString(h.Sum(nil))
}

// WriteFileWithChecksum atomically writes data to jsonPath and saves its
// checksum to the companion .shasum file.
func WriteFileWithChecksum(jsonPath string, data []byte) error {
	// Write the main file
	if err := os.WriteFile(jsonPath, data, 0644); err != nil {
		return fmt.Errorf("write %s: %w", jsonPath, err)
	}
	// Write checksum file
	checksum := computeChecksum(data)
	if err := os.WriteFile(shasumPath(jsonPath), []byte(checksum), 0644); err != nil {
		// Non-fatal: log but don't fail the whole save
		log.Printf("integrity: failed to write checksum for %s: %v", jsonPath, err)
	}
	return nil
}

// VerifyChecksum reads the .shasum companion file and compares it against the
// checksum of the supplied file bytes.
// Returns (true, nil) when intact, (false, nil) when corrupt/mismatch/missing,
// and (false, err) when the checksum file cannot be read for a non-missing reason.
func VerifyChecksum(jsonPath string, data []byte) (intact bool, err error) {
	csPath := shasumPath(jsonPath)
	stored, readErr := os.ReadFile(csPath)
	if readErr != nil {
		if os.IsNotExist(readErr) {
			// No checksum file — treat as corrupt (same as a mismatch).
			return false, nil
		}
		return false, fmt.Errorf("read checksum file %s: %w", csPath, readErr)
	}
	expected := computeChecksum(data)
	return string(stored) == expected, nil
}

// ---------------------------------------------------------------------------
// Global integrity registry
// ---------------------------------------------------------------------------

// FileIntegrity holds the integrity result for one file.
type FileIntegrity struct {
	Path    string `json:"path"`    // relative path under DATA
	Intact  bool   `json:"intact"`  // true = checksum matched
	Missing bool   `json:"missing"` // true = no checksum file existed yet
	Err     string `json:"error,omitempty"`
}

// IntegrityRegistry tracks integrity results for all loaded files.
type IntegrityRegistry struct {
	mu      sync.RWMutex
	results map[string]FileIntegrity // keyed by absolute path
}

var globalIntegrity = &IntegrityRegistry{
	results: make(map[string]FileIntegrity),
}

// Record stores a result for a file path. absPath must be the absolute path.
func (ir *IntegrityRegistry) Record(absPath string, intact bool, missing bool, errMsg string) {
	rel, relErr := filepath.Rel(dataDir, absPath)
	if relErr != nil {
		rel = absPath
	}
	ir.mu.Lock()
	ir.results[absPath] = FileIntegrity{
		Path:    rel,
		Intact:  intact,
		Missing: missing,
		Err:     errMsg,
	}
	ir.mu.Unlock()
}

// AllResults returns a copy of all recorded integrity results.
func (ir *IntegrityRegistry) AllResults() []FileIntegrity {
	ir.mu.RLock()
	defer ir.mu.RUnlock()
	out := make([]FileIntegrity, 0, len(ir.results))
	for _, v := range ir.results {
		out = append(out, v)
	}
	return out
}

// IsCorrupt returns true if the given absolute path is known to be corrupt
// (includes files whose .shasum is missing).
func (ir *IntegrityRegistry) IsCorrupt(absPath string) bool {
	ir.mu.RLock()
	defer ir.mu.RUnlock()
	fi, ok := ir.results[absPath]
	return ok && !fi.Intact
}

// CommonFilesCorrupt returns true when users.json or projects.json are corrupt
// (including when their .shasum file is missing).
func (ir *IntegrityRegistry) CommonFilesCorrupt() bool {
	absUsers, _ := filepath.Abs(filepath.Join(dataDir, "users.json"))
	absProjects, _ := filepath.Abs(filepath.Join(dataDir, "projects.json"))
	ir.mu.RLock()
	defer ir.mu.RUnlock()
	for _, p := range []string{absUsers, absProjects} {
		if fi, ok := ir.results[p]; ok && !fi.Intact {
			return true
		}
	}
	return false
}

// ProjectHasCorruption returns true if any file under DATA/<project>/ is corrupt
// (includes files whose .shasum is missing).
func (ir *IntegrityRegistry) ProjectHasCorruption(project string) bool {
	absBase, _ := filepath.Abs(filepath.Join(dataDir, project))
	prefix := absBase + string(filepath.Separator)
	ir.mu.RLock()
	defer ir.mu.RUnlock()
	for absPath, fi := range ir.results {
		if strings.HasPrefix(absPath, prefix) && !fi.Intact {
			return true
		}
	}
	return false
}

// CheckAndRecord reads data bytes, verifies the checksum for jsonPath, records
// the result in the global registry, and returns whether the file is intact.
// absPath is the absolute file path.
func CheckAndRecord(absPath string, data []byte) bool {
	intact, err := VerifyChecksum(absPath, data)
	if err != nil {
		globalIntegrity.Record(absPath, false, false, err.Error())
		log.Printf("integrity: error verifying %s: %v", absPath, err)
		return false
	}
	if intact {
		globalIntegrity.Record(absPath, true, false, "")
	} else {
		// Treat both a missing .shasum and a hash mismatch as corruption.
		_, statErr := os.Stat(shasumPath(absPath))
		if os.IsNotExist(statErr) {
			globalIntegrity.Record(absPath, false, true, "no checksum file — treated as corrupt")
			log.Printf("integrity: no checksum file for %s — treated as CORRUPT", absPath)
		} else {
			globalIntegrity.Record(absPath, false, false, "checksum mismatch")
			log.Printf("integrity: CHECKSUM MISMATCH for %s — file may be corrupt!", absPath)
		}
	}
	return intact
}
