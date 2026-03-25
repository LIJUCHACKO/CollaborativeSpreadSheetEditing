// serve/main.go
//
// Production static file server + reverse proxy for the shared-spreadsheet frontend.
//
// It replicates the proxy rules defined in vite.config.js:
//   - /api/**  → HTTP  → backend (default localhost:8082)
//   - /ws/**   → WS    → backend (default localhost:8082)
//   - everything else  → frontend dist/ directory (SPA fallback to index.html)
//
// Usage:
//
//	go run ./serve            # serves dist/ on :5175, backend on localhost:8082
//	./serve -listen :80 -backend 192.168.0.102:8082
//
// The -backend flag accepts either   host:port   or   http://host:port  .

package main

import (
	"flag"
	"log"
	"net/http"
	"net/http/httputil"
	"net/url"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	listen := flag.String("listen", ":5175", "address to listen on")
	backendFlag := flag.String("backend", "localhost:8082", "backend host:port (or http://host:port)")
	distDir := flag.String("dist", "", "path to frontend dist directory (default: <serve-binary-dir>/../dist)")
	flag.Parse()

	// ── resolve backend URL ──────────────────────────────────────────────────
	backendStr := *backendFlag
	if !strings.HasPrefix(backendStr, "http://") && !strings.HasPrefix(backendStr, "https://") {
		backendStr = "http://" + backendStr
	}
	backendHTTP, err := url.Parse(backendStr)
	if err != nil {
		log.Fatalf("invalid -backend URL %q: %v", *backendFlag, err)
	}
	backendWS, _ := url.Parse(backendStr)
	backendWS.Scheme = strings.Replace(backendWS.Scheme, "http", "ws", 1) // http→ws, https→wss

	// ── resolve dist directory ───────────────────────────────────────────────
	dist := *distDir
	if dist == "" {
		// Default: the dist/ folder that sits next to the serve/ directory
		exe, err := os.Executable()
		if err != nil {
			log.Fatalf("cannot determine executable path: %v", err)
		}
		dist = filepath.Join(filepath.Dir(exe), "..", "dist")
	}
	dist, err = filepath.Abs(dist)
	if err != nil {
		log.Fatalf("cannot resolve dist path: %v", err)
	}
	if _, err := os.Stat(dist); os.IsNotExist(err) {
		log.Fatalf("dist directory not found: %s", dist)
	}
	log.Printf("Serving frontend from : %s", dist)
	log.Printf("Backend (HTTP/API)    : %s", backendHTTP)
	log.Printf("Backend (WS)         : %s", backendWS)

	mux := http.NewServeMux()

	// ── /api  → HTTP reverse proxy ───────────────────────────────────────────
	apiProxy := httputil.NewSingleHostReverseProxy(backendHTTP)
	mux.HandleFunc("/api/", func(w http.ResponseWriter, r *http.Request) {
		r.Host = backendHTTP.Host
		apiProxy.ServeHTTP(w, r)
	})

	// ── /ws  → WebSocket reverse proxy ──────────────────────────────────────
	wsProxy := newWSProxy(backendWS)
	mux.HandleFunc("/ws", wsProxy)
	mux.HandleFunc("/ws/", wsProxy)

	// ── everything else → SPA static files ──────────────────────────────────
	mux.HandleFunc("/", spaHandler(dist))

	log.Printf("Listening on %s", *listen)
	if err := http.ListenAndServe(*listen, mux); err != nil {
		log.Fatal(err)
	}
}

// spaHandler serves files from distDir. When a path is not found it falls back
// to index.html so that client-side routing works correctly.
func spaHandler(distDir string) http.HandlerFunc {
	fileServer := http.FileServer(http.Dir(distDir))
	return func(w http.ResponseWriter, r *http.Request) {
		// Check whether the requested file exists inside dist/
		fsPath := filepath.Join(distDir, filepath.FromSlash(r.URL.Path))
		_, err := os.Stat(fsPath)
		if err == nil {
			// File exists – serve it directly (handles assets, js, css, etc.)
			fileServer.ServeHTTP(w, r)
			return
		}
		if !os.IsNotExist(err) {
			// Unexpected error
			http.Error(w, "internal error", http.StatusInternalServerError)
			return
		}
		// File not found – SPA fallback: serve index.html
		http.ServeFile(w, r, filepath.Join(distDir, "index.html"))
	}
}

// newWSProxy creates an http.HandlerFunc that tunnels WebSocket connections to
// the backend using raw TCP hijacking.  It rewrites the Host header and the
// target scheme/host, then copies bytes in both directions.
func newWSProxy(target *url.URL) http.HandlerFunc {
	proxy := httputil.NewSingleHostReverseProxy(target)

	// Override the Director so the Host header reaches the backend correctly.
	originalDirector := proxy.Director
	proxy.Director = func(req *http.Request) {
		originalDirector(req)
		req.Host = target.Host
	}

	return func(w http.ResponseWriter, r *http.Request) {
		// httputil.ReverseProxy handles WebSocket upgrade automatically
		// since Go 1.20 (it detects the Upgrade header and switches to
		// a plain TCP tunnel).  We just need to make sure the scheme is ws/wss.
		r2 := r.Clone(r.Context())
		r2.URL.Scheme = target.Scheme
		r2.URL.Host = target.Host
		r2.Host = target.Host
		proxy.ServeHTTP(w, r2)
	}
}
