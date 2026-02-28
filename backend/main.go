package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var addr = flag.String("addr", ":8082", "http service address")

// Global hub instance for WebSocket connections
var globalHub *Hub

func main() {
	flag.Parse()

	// Initialize Hub
	globalHub = newHub()
	go globalHub.run()

	globalProjectAuditManager.Load()
	globalProjectMeta.Load()
	// Initialize Sheet Manager (already initialized via global var in sheet.go, but good practice to be explicit if it wasn't)
	globalSheetManager.Load()
	globalUserManager.Load()
	globalChatManager.Load()

	// Start SheetManager async saver & flusher after Hub is ready
	// Ensures any broadcasts during script processing see a non-nil globalHub
	globalSheetManager.initAsyncSaver()

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

		sheetName := r.URL.Query().Get("sheet_name")
		if sheetName == "" {
			http.Error(w, "sheet_name is required", http.StatusBadRequest)
			return
		}
		project := r.URL.Query().Get("project")
		sheet := globalSheetManager.GetSheetBy(sheetName, project)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}

		f := excelize.NewFile()
		const xlsxSheetName = "Sheet1"
		f.NewSheet(xlsxSheetName)
		f.DeleteSheet("Sheet1")
		// Ensure we are working on a known sheet name
		f.NewSheet(xlsxSheetName)

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
				_ = f.SetCellValue(xlsxSheetName, cellRef, cell.Value)
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

		log.Printf("User %s exported sheet %s to XLSX", username, sheetName)
	})

	// Export all sheets in a project as XLSX
	http.HandleFunc("/api/export_project", func(w http.ResponseWriter, r *http.Request) {
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
		project := r.URL.Query().Get("project")
		if project == "" {
			http.Error(w, "project is required", http.StatusBadRequest)
			return
		}
		// Filter sheets by project (no ListSheetsByProject helper available)
		allSheets := globalSheetManager.ListSheets()
		sheets := make([]*Sheet, 0)
		for _, s := range allSheets {
			if s != nil && s.ProjectName == project {
				sheets = append(sheets, s)
			}
		}
		if len(sheets) == 0 {
			http.Error(w, "No sheets found for project", http.StatusNotFound)
			return
		}
		f := excelize.NewFile()
		for _, sheet := range sheets {
			sheetName := sheet.Name
			// Create sheet in workbook (ignore return values to avoid mismatch)
			f.NewSheet(sheetName)

			// Build column labels and max row based on data
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

			// Compute max row
			maxRow := 0
			for r := range rowSet {
				if r > maxRow {
					maxRow = r
				}
			}

			// Columns lexicographically sorted (A, B, ..., Z, AA, AB, ...)
			colLabels := make([]string, 0, len(colSet))
			for c := range colSet {
				colLabels = append(colLabels, c)
			}
			for i := 0; i < len(colLabels); i++ {
				for j := i + 1; j < len(colLabels); j++ {
					if colLabels[j] < colLabels[i] {
						colLabels[i], colLabels[j] = colLabels[j], colLabels[i]
					}
				}
			}

			// Write data rows
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
		}
		// Remove default sheet if present and unused
		_ = f.DeleteSheet("Sheet1")
		// Set active sheet to the first
		f.SetActiveSheet(1)
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		filename := project + "_" + time.Now().Format("20060102150405") + ".xlsx"
		w.Header().Set("Content-Disposition", "attachment; filename=\""+filename+"\"")
		if err := f.Write(w); err != nil {
			log.Printf("error writing project xlsx: %v", err)
			http.Error(w, "Failed to generate file", http.StatusInternalServerError)
			return
		}
		log.Printf("User %s exported project %s to XLSX", username, project)
	})

	// Import XLSX workbook into a project, creating sheets per workbook sheet
	http.HandleFunc("/api/import_project_xlsx", func(w http.ResponseWriter, r *http.Request) {
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

		project := r.URL.Query().Get("project")
		if project == "" {
			http.Error(w, "project is required", http.StatusBadRequest)
			return
		}

		// Parse multipart form to get the uploaded file
		if err := r.ParseMultipartForm(50 << 20); err != nil { // 50MB
			http.Error(w, "failed to parse form: "+err.Error(), http.StatusBadRequest)
			return
		}
		file, _, err := r.FormFile("file")
		if err != nil {
			http.Error(w, "file is required: "+err.Error(), http.StatusBadRequest)
			return
		}
		defer file.Close()

		// Open the workbook from the uploaded file stream
		f, err := excelize.OpenReader(file)
		if err != nil {
			http.Error(w, "failed to read xlsx: "+err.Error(), http.StatusBadRequest)
			return
		}
		defer func() { _ = f.Close() }()

		// Helpers to convert numeric column index to Excel label (A, B, ..., AA)
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

		created := make([]map[string]string, 0)
		sheetNames := f.GetSheetList()
		if len(sheetNames) == 0 {
			http.Error(w, "workbook has no sheets", http.StatusBadRequest)
			return
		}

		for _, wbSheetName := range sheetNames {
			rows, err := f.GetRows(wbSheetName)
			if err != nil {
				// Skip sheets that can't be read
				log.Printf("import: failed to get rows for sheet %s: %v", wbSheetName, err)
				continue
			}
			// Create a new sheet in project with same name
			newSheet := globalSheetManager.CreateSheet(wbSheetName, username, project, "datasheet")
			// Populate data
			newSheet.mu.Lock()
			if newSheet.Data == nil {
				newSheet.Data = make(map[string]map[string]Cell)
			}
			for rIdx, row := range rows {
				rowKey := strconv.Itoa(rIdx + 1) // 1-based
				if _, ok := newSheet.Data[rowKey]; !ok {
					newSheet.Data[rowKey] = make(map[string]Cell)
				}
				for cIdx, val := range row {
					if val == "" {
						continue
					}
					colLabel := toColLabel(cIdx + 1)
					newSheet.Data[rowKey][colLabel] = Cell{Value: val, User: username}
				}
			}
			newSheet.mu.Unlock()
			// Persist once
			globalSheetManager.SaveSheet(newSheet)
			fmt.Printf("Saving Sheet %s\n", newSheet.Name)
			time.Sleep(1000 * time.Millisecond) // slight delay to ensure filesystem consistency
			created = append(created, map[string]string{"id": newSheet.Name})
			// Project-level audit per sheet
			globalProjectAuditManager.Append(project, username, "IMPORT_SHEET", "Imported sheet '"+newSheet.Name+"' from XLSX")
		}

		if len(created) == 0 {
			http.Error(w, "no sheets imported", http.StatusBadRequest)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusCreated)
		json.NewEncoder(w).Encode(map[string]interface{}{"created": created})
		// Project-level audit summary
		globalProjectAuditManager.Append(project, username, "IMPORT_PROJECT_XLSX", "Imported "+strconv.Itoa(len(created))+" sheet(s) from uploaded XLSX")
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
		// Project-level audit: sheet copy (log on target and source)
		globalProjectAuditManager.Append(req.TargetProject, username, "COPY_SHEET", "Copied from project '"+req.SourceProject+"' sheet '"+req.SourceID+"' to '"+newSheet.Name+"'")
		if req.SourceProject != "" {
			globalProjectAuditManager.Append(req.SourceProject, username, "COPY_SHEET", "Copied sheet id="+req.SourceID+" to project '"+req.TargetProject+"' as id="+newSheet.Name)
		}
	})

	http.HandleFunc("/ws", func(w http.ResponseWriter, r *http.Request) {
		serveWs(globalHub, w, r)
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
		json.NewEncoder(w).Encode(map[string]interface{}{
			"token":              token,
			"username":           req.Username,
			"is_admin":           globalUserManager.IsAdminUser(req.Username),
			"can_create_project": globalUserManager.CanUserCreateProject(req.Username),
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

	// Change password for current user
	http.HandleFunc("/api/user/password", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		if r.Method == http.MethodOptions {
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
			OldPassword string `json:"old_password"`
			NewPassword string `json:"new_password"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.OldPassword == "" || req.NewPassword == "" {
			http.Error(w, "old_password and new_password are required", http.StatusBadRequest)
			return
		}

		if err := globalUserManager.ChangePassword(username, req.OldPassword, req.NewPassword); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "password updated"})
	})

	// ── Admin: GET /api/admin/users  (list all users)
	http.HandleFunc("/api/admin/users", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}
		token := r.Header.Get("Authorization")
		caller, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		if !globalUserManager.IsAdminUser(caller) {
			http.Error(w, "Forbidden", http.StatusForbidden)
			return
		}
		users := globalUserManager.ListUsers()
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(users)
	})

	// ── Admin: PUT /api/admin/user/password  (set any user's password)
	http.HandleFunc("/api/admin/user/password", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodPut {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		token := r.Header.Get("Authorization")
		caller, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		if !globalUserManager.IsAdminUser(caller) {
			http.Error(w, "Forbidden", http.StatusForbidden)
			return
		}
		var req struct {
			Username    string `json:"username"`
			NewPassword string `json:"new_password"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.Username == "" || req.NewPassword == "" {
			http.Error(w, "username and new_password are required", http.StatusBadRequest)
			return
		}
		if err := globalUserManager.AdminSetPassword(req.Username, req.NewPassword); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "password updated"})
	})

	// ── Admin: PUT /api/admin/user/permission  (grant/revoke project creation)
	http.HandleFunc("/api/admin/user/permission", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}
		if r.Method != http.MethodPut {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		token := r.Header.Get("Authorization")
		caller, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		if !globalUserManager.IsAdminUser(caller) {
			http.Error(w, "Forbidden", http.StatusForbidden)
			return
		}
		var req struct {
			Username         string `json:"username"`
			CanCreateProject bool   `json:"can_create_project"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.Username == "" {
			http.Error(w, "username is required", http.StatusBadRequest)
			return
		}
		if err := globalUserManager.SetCanCreateProject(req.Username, req.CanCreateProject); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "permission updated"})
	})

	// ── Admin: POST /api/admin/project/transfer  (change owner of a project)
	http.HandleFunc("/api/admin/project/transfer", func(w http.ResponseWriter, r *http.Request) {
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
		caller, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		if !globalUserManager.IsAdminUser(caller) {
			http.Error(w, "Forbidden", http.StatusForbidden)
			return
		}
		var req struct {
			Project  string `json:"project"`
			NewOwner string `json:"new_owner"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if strings.TrimSpace(req.Project) == "" || strings.TrimSpace(req.NewOwner) == "" {
			http.Error(w, "project and new_owner are required", http.StatusBadRequest)
			return
		}
		if !globalUserManager.Exists(req.NewOwner) {
			http.Error(w, "new owner does not exist", http.StatusBadRequest)
			return
		}

		// Update project meta owner
		globalProjectMeta.SetOwner(req.Project, req.NewOwner)
		// Append project-level audit entry
		globalProjectAuditManager.Append(req.Project, caller, "TRANSFER_PROJECT_OWNER", "Transferred project ownership to "+req.NewOwner)

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "project owner updated"})
	})

	// ── Admin: POST /api/admin/sheet/transfer  (change owner of a sheet)
	http.HandleFunc("/api/admin/sheet/transfer", func(w http.ResponseWriter, r *http.Request) {
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
		caller, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		if !globalUserManager.IsAdminUser(caller) {
			http.Error(w, "Forbidden", http.StatusForbidden)
			return
		}
		var req struct {
			Project   string `json:"project"`
			SheetName string `json:"sheet_name"`
			NewOwner  string `json:"new_owner"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if strings.TrimSpace(req.SheetName) == "" || strings.TrimSpace(req.NewOwner) == "" {
			http.Error(w, "sheet_name and new_owner are required", http.StatusBadRequest)
			return
		}
		if !globalUserManager.Exists(req.NewOwner) {
			http.Error(w, "new owner does not exist", http.StatusBadRequest)
			return
		}

		sheet := globalSheetManager.GetSheetBy(req.SheetName, req.Project)
		if sheet == nil {
			http.Error(w, "sheet not found", http.StatusNotFound)
			return
		}

		sheet.TransferOwnershipbyAdmin(req.NewOwner)
		globalSheetManager.SaveSheet(sheet)
		if req.Project != "" {
			globalProjectAuditManager.Append(req.Project, caller, "TRANSFER_SHEET_OWNER", "Transferred ownership of sheet '"+sheet.Name+"' to "+req.NewOwner)
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "sheet owner updated"})
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
			// Optional project filter via query parameter (e.g. /api/sheets?project=ProjectA)
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
			// Create new sheet in optional project (project specified in request body)
			var req struct {
				Name        string `json:"name"`
				User        string `json:"user"`
				ProjectName string `json:"project_name"`
				SheetType   string `json:"sheet_type"` // "datasheet" or "document"
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			// Reject if a sheet file or subfolder with the same name already exists
			{
				var dir string
				if req.ProjectName != "" {
					dir = filepath.Join(dataDir, req.ProjectName)
				} else {
					dir = dataDir
				}
				if _, statErr := os.Stat(filepath.Join(dir, req.Name+".json")); statErr == nil {
					http.Error(w, "A sheet with that name already exists", http.StatusConflict)
					return
				}
				if _, statErr := os.Stat(filepath.Join(dir, req.Name)); statErr == nil {
					http.Error(w, "A folder with that name already exists", http.StatusConflict)
					return
				}
			}
			// Only the project owner may create sheets inside a project/subfolder
			if req.ProjectName != "" {
				topProject := strings.SplitN(req.ProjectName, "/", 2)[0]
				owner := globalProjectMeta.GetOwner(topProject)
				if owner != "" && owner != username {
					http.Error(w, "Forbidden: only the project owner can create sheets here", http.StatusForbidden)
					return
				}
			}
			// Use authenticated username instead of client-provided user
			sheet := globalSheetManager.CreateSheet(req.Name, username, req.ProjectName, req.SheetType)
			// Project-level audit: sheet creation
			globalProjectAuditManager.Append(req.ProjectName, username, "CREATE_SHEET", "Created sheet '"+sheet.Name+"'")
			json.NewEncoder(w).Encode(sheet)
			return
		}

		if r.Method == "PUT" {
			// Rename a sheet - only owner may rename, and must not conflict with existing sheet or folder in the same project
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

			// Enforce owner-only rename
			s := globalSheetManager.GetSheetBy(req.ID, req.ProjectName)
			if s == nil {
				http.Error(w, "Sheet not found", http.StatusNotFound)
				return
			}
			if s.Owner != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			// Reject if a sheet file or subfolder with the new name already exists
			{
				var dir string
				if req.ProjectName != "" {
					dir = filepath.Join(dataDir, req.ProjectName)
				} else {
					dir = dataDir
				}
				if _, statErr := os.Stat(filepath.Join(dir, req.Name+".json")); statErr == nil {
					http.Error(w, "A sheet with that name already exists", http.StatusConflict)
					return
				}
				if _, statErr := os.Stat(filepath.Join(dir, req.Name)); statErr == nil {
					http.Error(w, "A folder with that name already exists", http.StatusConflict)
					return
				}
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
			// Only owner may delete, and must specify sheet ID and optional project in query parameters
			// Extract sheet ID from query parameter
			id := r.URL.Query().Get("id")
			if id == "" {
				http.Error(w, "Sheet ID required", http.StatusBadRequest)
				return
			}
			project := r.URL.Query().Get("project")
			// Fetch for audit details
			s := globalSheetManager.GetSheetBy(id, project)
			// Only owner may delete the sheet
			if s != nil && s.Owner != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			// Project-aware delete
			if !globalSheetManager.DeleteSheetBy(id, project) {
				http.Error(w, "Sheet not found", http.StatusNotFound)
				return
			}

			w.WriteHeader(http.StatusOK)
			json.NewEncoder(w).Encode(map[string]string{"message": "Sheet deleted"})
			// Project-level audit: sheet deletion
			if s != nil {
				globalProjectAuditManager.Append(project, username, "DELETE_SHEET", "Deleted sheet '"+s.Name+"'")
			} else {
				globalProjectAuditManager.Append(project, username, "DELETE_SHEET", "Deleted sheet id="+id)
			}
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
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		switch r.Method {
		case http.MethodGet:
			// List subdirectories in dataDir as projects
			entries, err := os.ReadDir(dataDir)
			if err != nil {
				http.Error(w, "Failed to read projects", http.StatusInternalServerError)
				return
			}
			type Project struct {
				Name  string `json:"name"`
				Owner string `json:"owner,omitempty"`
			}
			projects := make([]Project, 0)
			for _, e := range entries {
				if e.IsDir() {
					owner := globalProjectMeta.GetOwner(e.Name())
					projects = append(projects, Project{Name: e.Name(), Owner: owner})
				}
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(projects)
			return

		case http.MethodPost:
			// Create new project as a subdirectory under dataDir
			var req struct {
				Name string `json:"name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.Name == "" {
				http.Error(w, "Project name required", http.StatusBadRequest)
				return
			}
			// Check permission: only admin and approved users may create projects
			if !globalUserManager.CanUserCreateProject(username) {
				http.Error(w, "Not allowed: contact admin to get project creation permission", http.StatusForbidden)
				return
			}
			if _, statErr := os.Stat(filepath.Join(dataDir, req.Name)); statErr == nil {
				http.Error(w, "A project or folder with that name already exists", http.StatusConflict)
				return
			}
			if err := os.MkdirAll(filepath.Join(dataDir, req.Name), 0755); err != nil {
				http.Error(w, "Failed to create project", http.StatusInternalServerError)
				return
			}
			// Set project owner to the authenticated user
			globalProjectMeta.SetOwner(req.Name, username)
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"name": req.Name, "owner": username})
			return

		case http.MethodPut:
			// Rename a project (directory) - only owner may rename
			var req struct{ OldName, NewName string }
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.OldName == "" || req.NewName == "" {
				http.Error(w, "OldName and NewName required", http.StatusBadRequest)
				return
			}
			// Enforce owner-only rename
			if globalProjectMeta.GetOwner(req.OldName) != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			// Check if any sheets in the project or its subfolders are currently open
			if globalHub.HasActiveConnectionsForProject(req.OldName) {
				http.Error(w, "Cannot rename: one or more sheets in this project are currently open by users", http.StatusConflict)
				return
			}
			oldPath := filepath.Join(dataDir, req.OldName)
			newPath := filepath.Join(dataDir, req.NewName)
			if _, statErr := os.Stat(newPath); statErr == nil {
				http.Error(w, "A project or folder with that name already exists", http.StatusConflict)
				return
			}
			if err := os.Rename(oldPath, newPath); err != nil {
				http.Error(w, "Failed to rename project", http.StatusInternalServerError)
				return
			}
			// Preserve project owner mapping on rename
			globalProjectMeta.Rename(req.OldName, req.NewName)
			// Update in-memory sheets' ProjectName (including sheets in subfolders)
			for _, s := range globalSheetManager.ListSheets() {
				s.mu.RLock()
				oldPN := s.ProjectName
				s.mu.RUnlock()
				var newPN string
				if oldPN == req.OldName {
					newPN = req.NewName
				} else if strings.HasPrefix(oldPN, req.OldName+"/") {
					newPN = req.NewName + oldPN[len(req.OldName):]
				} else {
					continue
				}
				s.mu.Lock()
				s.ProjectName = newPN
				s.mu.Unlock()
				globalSheetManager.SaveSheet(s)
			}
			// Update script dependencies for the renamed project
			globalSheetManager.RenameProjectInDependencies(req.OldName, req.NewName)
			globalSheetManager.RenameProjectInOptionsRangeDependencies(req.OldName, req.NewName)
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Project renamed"})
			return

		case http.MethodDelete:
			// Extract project name from query parameter
			name := r.URL.Query().Get("name")
			if name == "" {
				http.Error(w, "Project name required", http.StatusBadRequest)
				return
			}
			// Only project owner may delete the project
			owner := globalProjectMeta.GetOwner(name)
			if owner == "" || owner != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			// Delete sheets in memory and files
			globalSheetManager.DeleteSheetsByProject(name)
			// Remove directory
			if err := os.RemoveAll(filepath.Join(dataDir, name)); err != nil {
				http.Error(w, "Failed to delete project", http.StatusInternalServerError)
				return
			}
			// Remove project ownership meta
			globalProjectMeta.Delete(name)
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Project deleted"})
			return

		default:
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
	})

	// Folders API: list/create subfolders under a project path
	http.HandleFunc("/api/folders", func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, PUT, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")

		token := r.Header.Get("Authorization")
		username, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		switch r.Method {
		case http.MethodGet:
			// Query immediate subfolders of the given project path
			projectPath := r.URL.Query().Get("project")
			if projectPath == "" {
				http.Error(w, "project is required", http.StatusBadRequest)
				return
			}
			abs := filepath.Join(dataDir, projectPath)
			entries, err := os.ReadDir(abs)
			if err != nil {
				http.Error(w, "Failed to read folders", http.StatusInternalServerError)
				return
			}
			type Folder struct {
				Name string `json:"name"`
			}
			folders := make([]Folder, 0)
			for _, e := range entries {
				if e.IsDir() {
					folders = append(folders, Folder{Name: e.Name()})
				}
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(folders)
			return

		case http.MethodPost:
			// Create a subfolder under the given parent project path
			var req struct {
				Parent string `json:"parent"`
				Name   string `json:"name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.Parent == "" || req.Name == "" {
				http.Error(w, "parent and name required", http.StatusBadRequest)
				return
			}
			// Only allow creation under an existing top-level project owned by user, if owner metadata exists
			top := strings.Split(req.Parent, string(os.PathSeparator))[0]
			owner := globalProjectMeta.GetOwner(top)
			if owner != "" && owner != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			abs := filepath.Join(dataDir, req.Parent, req.Name)
			if _, statErr := os.Stat(abs); statErr == nil {
				http.Error(w, "A folder or sheet with that name already exists", http.StatusConflict)
				return
			}
			if _, statErr := os.Stat(filepath.Join(dataDir, req.Parent, req.Name+".json")); statErr == nil {
				http.Error(w, "A sheet with that name already exists", http.StatusConflict)
				return
			}
			if err := os.MkdirAll(abs, 0755); err != nil {
				http.Error(w, "Failed to create folder", http.StatusInternalServerError)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"name": req.Name})
			return

		case http.MethodPut:
			// Rename a subfolder under the given parent project path
			var req struct {
				Parent  string `json:"parent"`
				OldName string `json:"old_name"`
				NewName string `json:"new_name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil || req.Parent == "" || req.OldName == "" || req.NewName == "" {
				http.Error(w, "parent, old_name, and new_name required", http.StatusBadRequest)
				return
			}
			// Only allow rename under an existing top-level project owned by user
			top := strings.Split(req.Parent, string(os.PathSeparator))[0]
			owner := globalProjectMeta.GetOwner(top)
			if owner != "" && owner != username {
				http.Error(w, "Forbidden: owner only", http.StatusForbidden)
				return
			}
			// Check if any sheets in the subfolder are currently open
			fullOldPath := req.Parent
			if fullOldPath != "" {
				fullOldPath = fullOldPath + "/" + req.OldName
			} else {
				fullOldPath = req.OldName
			}
			if globalHub.HasActiveConnectionsForProject(fullOldPath) {
				http.Error(w, "Cannot rename: one or more sheets in this folder are currently open by users", http.StatusConflict)
				return
			}
			oldPath := filepath.Join(dataDir, req.Parent, req.OldName)
			newPath := filepath.Join(dataDir, req.Parent, req.NewName)
			if _, statErr := os.Stat(newPath); statErr == nil {
				http.Error(w, "A folder or sheet with that name already exists", http.StatusConflict)
				return
			}
			if _, statErr := os.Stat(filepath.Join(dataDir, req.Parent, req.NewName+".json")); statErr == nil {
				http.Error(w, "A sheet with that name already exists", http.StatusConflict)
				return
			}
			if err := os.Rename(oldPath, newPath); err != nil {
				http.Error(w, "Failed to rename folder", http.StatusInternalServerError)
				return
			}
			// Update in-memory sheets' ProjectName for sheets in the renamed folder
			fullNewPath := req.Parent
			if fullNewPath != "" {
				fullNewPath = fullNewPath + "/" + req.NewName
			} else {
				fullNewPath = req.NewName
			}
			for _, s := range globalSheetManager.ListSheets() {
				if s.ProjectName == fullOldPath || strings.HasPrefix(s.ProjectName, fullOldPath+"/") {
					s.mu.Lock()
					// Replace the old prefix with the new prefix
					if s.ProjectName == fullOldPath {
						s.ProjectName = fullNewPath
					} else {
						s.ProjectName = fullNewPath + s.ProjectName[len(fullOldPath):]
					}
					s.mu.Unlock()
					globalSheetManager.SaveSheet(s)
				}
			}
			// Update script dependencies for the renamed subfolder
			globalSheetManager.RenameProjectInDependencies(fullOldPath, fullNewPath)
			globalSheetManager.RenameProjectInOptionsRangeDependencies(fullOldPath, fullNewPath)
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"name": req.NewName})
			return

		default:
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
	})

	// Copy-Paste a project or subfolder: copy all sheets to a new location with a new name
	http.HandleFunc("/api/projects/paste", func(w http.ResponseWriter, r *http.Request) {
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
			SourceType    string `json:"source_type"`     // "folder" (default) or "sheet"
			SourcePath    string `json:"source_path"`     // for folder: "project32io" or "project32io/sub1"; for sheet: project path
			SourceSheetID string `json:"source_sheet_id"` // only for source_type=sheet
			DestPath      string `json:"dest_path"`       // for folder: full dest path; for sheet: target project/folder path
			DestName      string `json:"dest_name"`       // only for source_type=sheet: new sheet name
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "Invalid request body", http.StatusBadRequest)
			return
		}
		if req.SourceType == "" {
			req.SourceType = "folder"
		}

		if req.SourceType == "sheet" {
			// --- Paste a single sheet into a target folder ---
			if req.SourceSheetID == "" || req.DestPath == "" {
				http.Error(w, "source_sheet_id and dest_path required for sheet paste", http.StatusBadRequest)
				return
			}
			newName := req.DestName
			if newName == "" {
				newName = req.SourceSheetID
			}
			// Ensure target folder exists on disk
			destDir := filepath.Join(dataDir, req.DestPath)
			if _, err := os.Stat(destDir); os.IsNotExist(err) {
				http.Error(w, "Destination folder not found", http.StatusNotFound)
				return
			}
			// Only the owner of the destination project may paste sheets there
			destTopProject := strings.SplitN(req.DestPath, "/", 2)[0]
			if destOwner := globalProjectMeta.GetOwner(destTopProject); destOwner != "" && destOwner != username {
				http.Error(w, "Forbidden: only the project owner can paste sheets here", http.StatusForbidden)
				return
			}
			newSheet := globalSheetManager.CopySheetToProject(req.SourceSheetID, req.SourcePath, req.DestPath, newName, username)
			if newSheet == nil {
				http.Error(w, "Source sheet not found", http.StatusNotFound)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "Sheet pasted successfully", "name": newSheet.Name})
			topProject := strings.SplitN(req.DestPath, "/", 2)[0]
			globalProjectAuditManager.Append(topProject, username, "PASTE_SHEET", "Pasted sheet '"+req.SourceSheetID+"' from '"+req.SourcePath+"' to '"+req.DestPath+"' as '"+newSheet.Name+"'")
			return
		}

		// --- Paste a folder/project ---
		if req.SourcePath == "" || req.DestPath == "" {
			http.Error(w, "source_path and dest_path required", http.StatusBadRequest)
			return
		}

		// Prevent pasting inside itself
		if req.DestPath == req.SourcePath || strings.HasPrefix(req.DestPath, req.SourcePath+"/") {
			http.Error(w, "Cannot paste a folder inside itself", http.StatusBadRequest)
			return
		}
		// Only the owner of the destination top-level project may paste inside it
		destTopParts := strings.SplitN(req.DestPath, "/", 2)
		if len(destTopParts) > 1 {
			// Pasting inside an existing project – check owner
			if destOwner := globalProjectMeta.GetOwner(destTopParts[0]); destOwner != "" && destOwner != username {
				http.Error(w, "Forbidden: only the project owner can paste here", http.StatusForbidden)
				return
			}
		} else {
			// Pasting as a new top-level project – requires create-project permission
			if !globalUserManager.CanUserCreateProject(username) {
				http.Error(w, "Not allowed: contact admin to get project creation permission", http.StatusForbidden)
				return
			}
		}

		// Ensure source exists
		if _, err := os.Stat(filepath.Join(dataDir, req.SourcePath)); os.IsNotExist(err) {
			http.Error(w, "Source path not found", http.StatusNotFound)
			return
		}
		// Ensure destination doesn't exist
		if _, err := os.Stat(filepath.Join(dataDir, req.DestPath)); err == nil {
			http.Error(w, "Destination already exists", http.StatusConflict)
			return
		}

		// Copy the folder structure on disk (empty folders that may not have sheets)
		srcAbs := filepath.Join(dataDir, req.SourcePath)
		destAbs := filepath.Join(dataDir, req.DestPath)
		if err := copyDirStructure(srcAbs, destAbs); err != nil {
			log.Printf("copy dir structure error: %v", err)
		}

		// If this is a top-level project paste, set owner
		destParts := strings.SplitN(req.DestPath, "/", 2)
		destTopProject := destParts[0]
		if len(destParts) == 1 {
			// Top-level project paste - set owner
			globalProjectMeta.SetOwner(destTopProject, username)
		}

		if err := globalSheetManager.CopyPasteProject(req.SourcePath, req.DestPath, username); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]string{"message": "Pasted successfully"})
		// Audit
		topProject := strings.SplitN(req.DestPath, "/", 2)[0]
		globalProjectAuditManager.Append(topProject, username, "PASTE_PROJECT", "Pasted from '"+req.SourcePath+"' to '"+req.DestPath+"'")
	})

	// Project audit API: list audit entries for a project
	http.HandleFunc("/api/projects/audit", func(w http.ResponseWriter, r *http.Request) {
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
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}
		project := r.URL.Query().Get("project")
		if project == "" {
			http.Error(w, "project is required", http.StatusBadRequest)
			return
		}
		entries := globalProjectAuditManager.List(project)
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(entries)
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
		json.NewEncoder(w).Encode(sheet.SnapshotForClient())
	})

	// Get the value of a specific cell in a sheet
	//usage
	//curl -s -H "Authorization: <token>" \
	// "http://localhost:8082/api/sheet/cell?sheet_name=20260122223748&project=project32/new2&row=1&col=A"
	//curl -s -H "Authorization: <token>" \
	// "http://localhost:8082/api/sheet/cell?sheet_name=20260122223748&project=project32/new2&cell=A1"
	//# With jq (recommended)
	//TOKEN=$(curl -s -X POST http://localhost:8082/api/login \
	// -H "Content-Type: application/json" \
	//  -d '{"username":"alice","password":"secret"}' | jq -r '.token')
	// # Without jq (basic grep/sed fallback)
	//TOKEN=$(curl -s -X POST http://localhost:8082/api/login \
	// -H "Content-Type: application/json" \
	//  -d '{"username":"alice","password":"secret"}' | sed -n 's/.*"token":"\([^"]*\)".*/\1/p')
	//this function is not tested
	http.HandleFunc("/api/sheet/cell", func(w http.ResponseWriter, r *http.Request) {
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
		_, err := globalUserManager.ValidateToken(token)
		if err != nil {
			http.Error(w, "Unauthorized: "+err.Error(), http.StatusUnauthorized)
			return
		}

		sheetName := r.URL.Query().Get("sheet_name")
		project := r.URL.Query().Get("project")
		row := r.URL.Query().Get("row")
		col := r.URL.Query().Get("col")
		cell := strings.TrimSpace(r.URL.Query().Get("cell"))

		if sheetName == "" {
			http.Error(w, "sheet_name is required", http.StatusBadRequest)
			return
		}

		// Allow either row+col or a combined cell like A5
		if cell != "" {
			// Parse leading letters as column and trailing digits as row
			i := 0
			for i < len(cell) {
				ch := cell[i]
				if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
					i++
					continue
				}
				break
			}
			colPart := cell[:i]
			rowPart := cell[i:]
			if colPart == "" || rowPart == "" {
				http.Error(w, "invalid cell format; expected like A5", http.StatusBadRequest)
				return
			}
			// Validate row is numeric
			if _, err := strconv.Atoi(rowPart); err != nil {
				http.Error(w, "invalid row number in cell", http.StatusBadRequest)
				return
			}
			col = strings.ToUpper(colPart)
			row = rowPart
		}

		if row == "" || col == "" {
			http.Error(w, "row and col (or cell) required", http.StatusBadRequest)
			return
		}
		sheet := globalSheetManager.GetSheetBy(sheetName, project)
		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}

		var value string
		exists := false
		sheet.mu.RLock()
		if rowMap, ok := sheet.Data[row]; ok {
			if cell, ok := rowMap[col]; ok {
				value = cell.Value
				exists = true
			}
		}
		sheet.mu.RUnlock()

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]interface{}{
			"sheet_name": sheetName,
			"project":    project,
			"row":        row,
			"col":        col,
			"value":      value,
			"exists":     exists,
		})
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

	// User preferences: visible rows/cols (common across sheets/projects)
	http.HandleFunc("/api/user/preferences", func(w http.ResponseWriter, r *http.Request) {
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

		switch r.Method {
		case http.MethodGet:
			prefs, err := globalUserManager.GetPreferences(username)
			if err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(prefs)
		case http.MethodPut:
			var req struct {
				VisibleRows int `json:"visible_rows"`
				VisibleCols int `json:"visible_cols"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			if err := globalUserManager.UpdatePreferences(username, Preferences{VisibleRows: req.VisibleRows, VisibleCols: req.VisibleCols}); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(map[string]string{"message": "preferences updated"})
		default:
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		}
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

		sheetName := r.URL.Query().Get("sheet_name")
		if sheetName == "" {
			http.Error(w, "sheet_name is required", http.StatusBadRequest)
			return
		}
		project := r.URL.Query().Get("project")
		sheet := globalSheetManager.GetSheetBy(sheetName, project)
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
			isAdmin := globalUserManager.IsAdminUser(username)
			if !sheet.UpdatePermissions(req.Editors, username, isAdmin) {
				http.Error(w, "Forbidden: owner or admin only", http.StatusForbidden)
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
			SheetName   string `json:"sheet_name"`
			NewOwner    string `json:"new_owner"`
			ProjectName string `json:"project_name"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		if req.SheetName == "" || req.NewOwner == "" {
			http.Error(w, "sheet_name and new_owner required", http.StatusBadRequest)
			return
		}
		isAdmin := globalUserManager.IsAdminUser(username)
		sheet := globalSheetManager.GetSheetBy(req.SheetName, req.ProjectName)

		if sheet == nil {
			http.Error(w, "Sheet not found", http.StatusNotFound)
			return
		}
		if !globalUserManager.Exists(req.NewOwner) {
			http.Error(w, "new_owner does not exist", http.StatusBadRequest)
			return
		}
		if !sheet.TransferOwnership(req.NewOwner, username, isAdmin) {
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

// copyDirStructure recursively copies the directory tree from src to dst,
// creating all subdirectories. It does NOT copy files (sheets are handled by CopyPasteProject).
func copyDirStructure(src, dst string) error {
	return filepath.Walk(src, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() {
			return nil // skip files, only create directories
		}
		rel, err := filepath.Rel(src, path)
		if err != nil {
			return err
		}
		target := filepath.Join(dst, rel)
		return os.MkdirAll(target, 0755)
	})
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
