package main

import (
	"encoding/json"
	"flag"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

var addr = flag.String("addr", ":8080", "http service address")

func main() {
	flag.Parse()

	// Initialize Hub
	hub := newHub()
	go hub.run()

	// Initialize Sheet Manager (already initialized via global var in sheet.go, but good practice to be explicit if it wasn't)
	globalSheetManager.Load()
	globalUserManager.Load()
	globalChatManager.Load()

	http.HandleFunc("/api/export", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}

		if r.Method != http.MethodGet {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		sheetID := r.URL.Query().Get("sheet_id")
		if sheetID == "" {
			http.Error(w, "sheet_id is required", http.StatusBadRequest)
			return
		}
		project := r.URL.Query().Get("project")
		sheet := globalSheetManager.GetSheetBy(sheetID, project)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}

		f := excelize.NewFile()
		const sheetName = "Sheet1"
		f.NewSheet(sheetName)
		f.DeleteSheet("Sheet1")
		// Ensure we are working on a known sheet name
		f.NewSheet(sheetName)

		// Write header row based on columns present in data
		colSet := make(map[string]struct{})
		rowSet := make(map[int]struct{})

		sheet.mu.RLock()
		for rowKey, cols := range sheet.Data {
			rowNum, _ := strconv.Atoi(rowKey)
			rowSet[rowNum] = struct{}{}
			for colKey := range cols {
				colSet[colKey] = struct{}{}
			}
		}

		// Collect and sort rows and columns
		maxRow := 0
		for r := range rowSet {
			if r > maxRow {
				maxRow = r
			}
		}

		// Columns: use existing labels sorted lexicographically
		colLabels := make([]string, 0, len(colSet))
		for c := range colSet {
			colLabels = append(colLabels, c)
		}

		// Simple lexicographic sort (A, B, ..., Z, AA, AB, ...)
		for i := 0; i < len(colLabels); i++ {
			for j := i + 1; j < len(colLabels); j++ {
				if colLabels[j] < colLabels[i] {
					colLabels[i], colLabels[j] = colLabels[j], colLabels[i]
				}
			}
		}

		// Write header row (row 1) with column labels
		/*
			for i, colLabel := range colLabels {
				cellRef, _ := excelize.CoordinatesToCellName(i+1, 1)
				_ = f.SetCellValue(sheetName, cellRef, colLabel)
			}
		*/
		// Write data starting from row 2
		for row := 1; row <= maxRow; row++ {
			rowKey := strconv.Itoa(row)
			cols, ok := sheet.Data[rowKey]
			if !ok {
				continue
			}
			for i, colLabel := range colLabels {
				cell, ok := cols[colLabel]
				if !ok {
					continue
				}
				cellRef, _ := excelize.CoordinatesToCellName(i+1, row)
				_ = f.SetCellValue(sheetName, cellRef, cell.Value)
			}
		}
		sheet.mu.RUnlock()

		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		filename := sheet.Name + "_" + time.Now().Format("20060102150405") + ".xlsx"
		w.Header().Set("Content-Disposition", "attachment; filename=\""+filename+"\"")

		if err := f.Write(w); err != nil {
			log.Printf("error writing xlsx: %v", err)
			http.Error(w, "Failed to generate file", http.StatusInternalServerError)
			return
		}

		log.Printf("User %s exported sheet %s to XLSX", username, sheetID)
	})

	// Copy a sheet from one project to another
	http.HandleFunc("/api/sheet/copy", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		var req struct {
			SourceID      string `json:"source_id"`
			SourceProject string `json:"source_project"`
			TargetProject string `json:"target_project"`
			Name          string `json:"name"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.SourceID == "" || req.TargetProject == "" {
			http.Error(w, "source_id and target_project required", http.StatusBadRequest)
			return
		}

		newSheet := globalSheetManager.CopySheetToProject(req.SourceID, req.SourceProject, req.TargetProject, req.Name, username)
		if newSheet == nil {
			http.Error(w, "Source sheet not found", http.StatusNotFound)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusCreated)
		json.NewEncoder(w).Encode(newSheet)
	})

	http.HandleFunc("/ws", func(w http.ResponseWriter, r *http.Request) {
		serveWs(hub, w, r)
	})

	http.HandleFunc("/api/register", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		if r.Method != "POST" {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		var req struct {
			Username string `json:"username"`
			Password string `json:"password"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		if req.Username == "" || req.Password == "" {
			http.Error(w, "Username and password required", http.StatusBadRequest)
			return
		}

		if err := globalUserManager.Register(req.Username, req.Password); err != nil {
			http.Error(w, err.Error(), http.StatusConflict) // User exists
			return
		}

		w.WriteHeader(http.StatusCreated)
	})

	http.HandleFunc("/api/login", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		if r.Method != "POST" {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		var req struct {
			Username string `json:"username"`
			Password string `json:"password"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		token, err := globalUserManager.Login(req.Username, req.Password)
		if err != nil {
			http.Error(w, err.Error(), http.StatusUnauthorized)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{
			"token":    token,
			"username": req.Username,
		})
	})

	http.HandleFunc("/api/logout", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		if r.Method != "POST" {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		token := r.Header.Get("Authorization")
		if token != "" {
			globalUserManager.Logout(token)
		}

		w.WriteHeader(http.StatusOK)
		json.NewEncoder(w).Encode(map[string]string{"message": "Logged out successfully"})
	})

	http.HandleFunc("/api/validate", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		token := r.Header.Get("Authorization")
		if token == "" {
			http.Error(w, "No token provided", http.StatusUnauthorized)
			return
		}

		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, err.Error(), http.StatusUnauthorized)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{
			"username": username,
			"valid":    "true",
		})
	})

	http.HandleFunc("/api/sheets", func(w http.ResponseWriter, r *http.Request) {
		// Simple CORS
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, DELETE, PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		// Validate token for protected endpoints
		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		if r.Method == "GET" {
			project := r.URL.Query().Get("project")
			all := globalSheetManager.ListSheets()
			if project == "" {
				json.NewEncoder(w).Encode(all)
				return
			}
			filtered := make([]*Sheet, 0)
			for _, s := range all {
				if s != nil && s.ProjectName == project {
					filtered = append(filtered, s)
				}
			}
			json.NewEncoder(w).Encode(filtered)
			return
		}

		if r.Method == "POST" {
			var req struct {
				Name        string `json:"name"`
				User        string `json:"user"`
				ProjectName string `json:"project_name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			// Use authenticated username instead of client-provided user
			sheet := globalSheetManager.CreateSheet(req.Name, username, req.ProjectName)
			json.NewEncoder(w).Encode(sheet)
			return
		}

		if r.Method == "PUT" {
			var req struct {
				ID          string `json:"id"`
				Name        string `json:"name"`
				ProjectName string `json:"project_name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}

			if req.ID == "" || req.Name == "" {
				http.Error(w, "Sheet ID and name required", http.StatusBadRequest)
				return
			}

			if !globalSheetManager.RenameSheetBy(req.ID, req.ProjectName, req.Name, username) {
				http.Error(w, "Sheet not found", http.StatusNotFound)
				return
			}

			w.WriteHeader(http.StatusOK)
			json.NewEncoder(w).Encode(map[string]string{"message": "Sheet renamed successfully"})
			return
		}

		if r.Method == "DELETE" {
			// Extract sheet ID from query parameter
			id := r.URL.Query().Get("id")
			if id == "" {
				http.Error(w, "Sheet ID required", http.StatusBadRequest)
				return
			}
			project := r.URL.Query().Get("project")
			// Project-aware delete
			if !globalSheetManager.DeleteSheetBy(id, project) {
				http.Error(w, "Sheet not found", http.StatusNotFound)
				return
			}

			w.WriteHeader(http.StatusOK)
			json.NewEncoder(w).Encode(map[string]string{"message": "Sheet deleted"})
			return
		}
	})

	// Projects API (filesystem-backed via DATA/<project>)
	http.HandleFunc("/api/projects", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}

		token := r.Header.Get("Authorization")
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		switch r.Method {
		case http.MethodGet:
			entries, err := os.ReadDir(dataDir)
			if err != nil {
				http.Error(w, "Failed to read projects", http.StatusInternalServerError)
				return
			}
			type Project struct {
				Name string `json:"name"`
			}
			projects := make([]Project, 0)
			for _, e := range entries {
				if e.IsDir() {
					projects = append(projects, Project{Name: e.Name()})
				}
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(projects)
			return

		case http.MethodPost:
			var req struct {
				Name string `json:"name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.Name == "" {
				http.Error(w, "Project name required", http.StatusBadRequest)
				return
			}
			if err := os.MkdirAll(filepath.Join(dataDir, req.Name), 0755); err != nil {
				http.Error(w, "Failed to create project", http.StatusInternalServerError)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"name": req.Name})
			return

		case http.MethodPut:
			var req struct{ OldName, NewName string }
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.OldName == "" || req.NewName == "" {
				http.Error(w, "OldName and NewName required", http.StatusBadRequest)
				return
			}
			oldPath := filepath.Join(dataDir, req.OldName)
			newPath := filepath.Join(dataDir, req.NewName)
			if err := os.Rename(oldPath, newPath); err != nil {
				http.Error(w, "Failed to rename project", http.StatusInternalServerError)
				return
			}
			// Update in-memory sheets' ProjectName
			for _, s := range globalSheetManager.ListSheets() {
				if s.ProjectName == req.OldName {
					s.mu.Lock()
					s.ProjectName = req.NewName
					s.mu.Unlock()
					globalSheetManager.SaveSheet(s)
				}
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Project renamed"})
			return

		case http.MethodDelete:
			name := r.URL.Query().Get("name")
			if name == "" {
				http.Error(w, "Project name required", http.StatusBadRequest)
				return
			}
			// Delete sheets in memory and files
			globalSheetManager.DeleteSheetsByProject(name)
			// Remove directory
			if err := os.RemoveAll(filepath.Join(dataDir, name)); err != nil {
				http.Error(w, "Failed to delete project", http.StatusInternalServerError)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Project deleted"})
			return

		default:
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
	})

	// Duplicate a project: copy all sheets to a new project
	http.HandleFunc("/api/projects/duplicate", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		token := r.Header.Get("Authorization")
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		var req struct {
			Source string `json:"source_name"`
			New    string `json:"new_name"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.Source == "" || req.New == "" {
			http.Error(w, "source_name and new_name required", http.StatusBadRequest)
			return
		}

		// Ensure source exists
		if _, err := os.Stat(filepath.Join(dataDir, req.Source)); os.IsNotExist(err) {
			http.Error(w, "Source project not found", http.StatusNotFound)
			return
		}
		// Ensure destination doesn't exist
		if _, err := os.Stat(filepath.Join(dataDir, req.New)); err == nil {
			http.Error(w, "Destination project already exists", http.StatusConflict)
			return
		}

		if err := globalSheetManager.DuplicateProject(req.Source, req.New); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "Project duplicated"})
	})

	// Get a single sheet by id
	http.HandleFunc("/api/sheet", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		if r.Method != http.MethodGet {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		token := r.Header.Get("Authorization")
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		id := r.URL.Query().Get("id")
		if id == "" {
			http.Error(w, "id is required", http.StatusBadRequest)
			return
		}
		project := r.URL.Query().Get("project")
		sheet := globalSheetManager.GetSheetBy(id, project)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(sheet)
	})

	// List all usernames (for selection)
	http.HandleFunc("/api/users", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodGet {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		token := r.Header.Get("Authorization")
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}
		users := globalUserManager.ListUsernames()
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(users)
	})

	// Get/Update permissions for a sheet
	http.HandleFunc("/api/sheet/permissions", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}

		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		sheetID := r.URL.Query().Get("sheet_id")
		if sheetID == "" {
			http.Error(w, "sheet_id is required", http.StatusBadRequest)
			return
		}
		project := r.URL.Query().Get("project")
		sheet := globalSheetManager.GetSheetBy(sheetID, project)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}

		if r.Method == http.MethodGet {
			resp := map[string]interface{}{
				"owner":       sheet.Owner,
				"permissions": sheet.Permissions,
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(resp)
			return
		}

		if r.Method == http.MethodPut {
			var req struct {
				Editors []string `json:"editors"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			if !sheet.UpdatePermissions(req.Editors, username) {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Permissions updated"})
			return
		}

		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
	})

	// Transfer ownership of a sheet
	http.HandleFunc("/api/sheet/transfer_owner", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodPut {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}
		var req struct {
			SheetID     string `json:"sheet_id"`
			NewOwner    string `json:"new_owner"`
			ProjectName string `json:"project_name"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.SheetID == "" || req.NewOwner == "" {
			http.Error(w, "sheet_id and new_owner required", http.StatusBadRequest)
			return
		}
		sheet := globalSheetManager.GetSheetBy(req.SheetID, req.ProjectName)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}
		if !globalUserManager.Exists(req.NewOwner) {
			http.Error(w, "new_owner does not exist", http.StatusBadRequest)
			return
		}
		if !sheet.TransferOwnership(req.NewOwner, username) {
			http.Error(w, "Forbidden: owner only", http.StatusForbidden)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "Ownership transferred"})
	})

	// Simple health check
	http.HandleFunc("/health", func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte("OK"))
	})

	log.Printf("Server started on %s", *addr)
	// Wrap DefaultServeMux with a global CORS middleware so that even 404/405 responses
	// include the appropriate CORS headers. This prevents CORS failures on project
	// duplication and sheet copy requests when paths/methods mismatch or errors occur.
	err := http.ListenAndServe(*addr, corsMiddleware(http.DefaultServeMux))
	if err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}

// corsMiddleware ensures CORS headers are present on every response, including errors
// and non-matching routes. It also short-circuits OPTIONS preflight requests.
func corsMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}

		next.ServeHTTP(w, r)
	})
}
