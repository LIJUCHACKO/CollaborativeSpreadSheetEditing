package main

import (
	"encoding/json"
	"flag"
	"log"
	"net/http"
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

		sheet := globalSheetManager.GetSheet(sheetID)
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
			// Ensure we load if not loaded? Or assume Load called at start.
			// Let's assume Load called at main.
			sheets := globalSheetManager.ListSheets()
			json.NewEncoder(w).Encode(sheets)
			return
		}

		if r.Method == "POST" {
			var req struct {
				Name string `json:"name"`
				User string `json:"user"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}
			// Use authenticated username instead of client-provided user
			sheet := globalSheetManager.CreateSheet(req.Name, username)
			json.NewEncoder(w).Encode(sheet)
			return
		}

		if r.Method == "PUT" {
			var req struct {
				ID   string `json:"id"`
				Name string `json:"name"`
			}
			if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
				http.Error(w, err.Error(), http.StatusBadRequest)
				return
			}

			if req.ID == "" || req.Name == "" {
				http.Error(w, "Sheet ID and name required", http.StatusBadRequest)
				return
			}

			if !globalSheetManager.RenameSheet(req.ID, req.Name, username) {
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

			if !globalSheetManager.DeleteSheet(id) {
				http.Error(w, "Sheet not found", http.StatusNotFound)
				return
			}

			w.WriteHeader(http.StatusOK)
			json.NewEncoder(w).Encode(map[string]string{"message": "Sheet deleted"})
			return
		}
	})

	// Simple health check
	http.HandleFunc("/health", func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte("OK"))
	})

	log.Printf("Server started on %s", *addr)
	err := http.ListenAndServe(*addr, nil)
	if err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}
