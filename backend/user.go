package main

import (
	"crypto/rand"
	"encoding/base64"
	"encoding/json"
	"errors"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"golang.org/x/crypto/bcrypt"
)

var userPersistenceFile = filepath.Join(dataDir, "users.json")

const sessionTimeout = 1 * time.Hour

type User struct {
	Username         string      `json:"username"`
	PasswordHash     string      `json:"password_hash"`
	Prefs            Preferences `json:"prefs,omitempty"`
	IsAdmin          bool        `json:"is_admin,omitempty"`
	CanCreateProject bool        `json:"can_create_project,omitempty"`
}

type Session struct {
	Token     string
	Username  string
	ExpiresAt time.Time
}

// Preferences holds user-level settings common across sheets/projects
type Preferences struct {
	VisibleRows int `json:"visible_rows,omitempty"`
	VisibleCols int `json:"visible_cols,omitempty"`
}

type UserManager struct {
	users    map[string]*User
	sessions map[string]*Session // token -> Session
	mu       sync.RWMutex
}

var globalUserManager = &UserManager{
	users:    make(map[string]*User),
	sessions: make(map[string]*Session),
}

func (um *UserManager) Register(username, password string) error {
	um.mu.Lock()
	defer um.mu.Unlock()

	// Disallow reserved usernames "system" and "admin" (case-insensitive)
	trimmed := strings.TrimSpace(username)
	if strings.EqualFold(trimmed, "system") || strings.EqualFold(trimmed, "admin") {
		return errors.New("reserved username")
	}

	if _, exists := um.users[username]; exists {
		return errors.New("user already exists")
	}

	hashedPassword, err := bcrypt.GenerateFromPassword([]byte(password), bcrypt.DefaultCost)
	if err != nil {
		return err
	}

	user := &User{
		Username:         username,
		PasswordHash:     string(hashedPassword),
		Prefs:            Preferences{VisibleRows: 15, VisibleCols: 7},
		CanCreateProject: false, // must be approved by admin
	}

	um.users[username] = user
	um.saveUsersLocked()
	return nil
}

func (um *UserManager) Login(username, password string) (string, error) {
	um.mu.RLock()
	user, exists := um.users[username]
	um.mu.RUnlock()

	if !exists {
		return "", errors.New("invalid credentials")
	}

	err := bcrypt.CompareHashAndPassword([]byte(user.PasswordHash), []byte(password))
	if err != nil {
		return "", errors.New("invalid credentials")
	}

	// Generate session token
	token, err := generateToken()
	if err != nil {
		return "", errors.New("failed to generate session token")
	}

	session := &Session{
		Token:     token,
		Username:  username,
		ExpiresAt: time.Now().Add(sessionTimeout),
	}

	um.mu.Lock()
	um.sessions[token] = session
	um.mu.Unlock()

	// Start cleanup goroutine for expired sessions
	go um.cleanupExpiredSessions()

	return token, nil
}

// generateToken creates a secure random token
func generateToken() (string, error) {
	b := make([]byte, 32)
	if _, err := rand.Read(b); err != nil {
		return "", err
	}
	return base64.URLEncoding.EncodeToString(b), nil
}

// ValidateToken checks if a token is valid and not expired
func (um *UserManager) ValidateToken(token string) (string, error) {
	um.mu.RLock()
	session, exists := um.sessions[token]
	um.mu.RUnlock()

	if !exists {
		return "", errors.New("invalid token")
	}

	if time.Now().After(session.ExpiresAt) {
		um.mu.Lock()
		delete(um.sessions, token)
		um.mu.Unlock()
		return "", errors.New("session expired")
	}

	return session.Username, nil
}

// Logout removes a session token
func (um *UserManager) Logout(token string) {
	um.mu.Lock()
	defer um.mu.Unlock()
	delete(um.sessions, token)
}

// cleanupExpiredSessions removes expired sessions periodically
func (um *UserManager) cleanupExpiredSessions() {
	um.mu.Lock()
	defer um.mu.Unlock()

	now := time.Now()
	for token, session := range um.sessions {
		if now.After(session.ExpiresAt) {
			delete(um.sessions, token)
		}
	}
}

// Exists returns true if a user with the given username exists
func (um *UserManager) Exists(username string) bool {
	um.mu.RLock()
	defer um.mu.RUnlock()
	_, ok := um.users[username]
	return ok
}

// ListUsernames returns a list of all registered usernames
func (um *UserManager) ListUsernames() []string {
	um.mu.RLock()
	defer um.mu.RUnlock()
	list := make([]string, 0, len(um.users))
	for uname := range um.users {
		list = append(list, uname)
	}
	return list
}

// IsAdminUser returns true if the given username is an admin
func (um *UserManager) IsAdminUser(username string) bool {
	um.mu.RLock()
	defer um.mu.RUnlock()
	u, ok := um.users[username]
	return ok && u.IsAdmin
}

// CanCreateProject returns true if user is allowed to create projects
func (um *UserManager) CanUserCreateProject(username string) bool {
	um.mu.RLock()
	defer um.mu.RUnlock()
	u, ok := um.users[username]
	return ok && (u.IsAdmin || u.CanCreateProject)
}

// SetCanCreateProject grants or revokes project-creation permission for a user
func (um *UserManager) SetCanCreateProject(username string, allowed bool) error {
	um.mu.Lock()
	defer um.mu.Unlock()
	u, ok := um.users[username]
	if !ok {
		return errors.New("user not found")
	}
	u.CanCreateProject = allowed
	um.saveUsersLocked()
	return nil
}

// AdminSetPassword forcefully sets a new password for any user (admin action, no old-password check)
func (um *UserManager) AdminSetPassword(username, newPassword string) error {
	um.mu.Lock()
	defer um.mu.Unlock()
	u, ok := um.users[username]
	if !ok {
		return errors.New("user not found")
	}
	if len(newPassword) < 6 {
		return errors.New("new password must be at least 6 characters")
	}
	hashedPassword, err := bcrypt.GenerateFromPassword([]byte(newPassword), bcrypt.DefaultCost)
	if err != nil {
		return err
	}
	u.PasswordHash = string(hashedPassword)
	um.saveUsersLocked()
	return nil
}

// ListUsers returns info about all users (safe for admin view)
type UserInfo struct {
	Username         string `json:"username"`
	IsAdmin          bool   `json:"is_admin"`
	CanCreateProject bool   `json:"can_create_project"`
}

func (um *UserManager) ListUsers() []UserInfo {
	um.mu.RLock()
	defer um.mu.RUnlock()
	list := make([]UserInfo, 0, len(um.users))
	for _, u := range um.users {
		list = append(list, UserInfo{
			Username:         u.Username,
			IsAdmin:          u.IsAdmin,
			CanCreateProject: u.IsAdmin || u.CanCreateProject,
		})
	}
	return list
}

func (um *UserManager) Load() {
	um.mu.Lock()
	defer um.mu.Unlock()

	file, err := os.Open(userPersistenceFile)
	if err != nil {
		if os.IsNotExist(err) {
			um.ensureAdminLocked()
			return
		}
		log.Printf("Error opening users file: %v", err)
		return
	}
	defer file.Close()

	var loadedUsers map[string]*User
	if err := json.NewDecoder(file).Decode(&loadedUsers); err != nil {
		log.Printf("Error decoding users: %v", err)
		return
	}

	um.users = loadedUsers
	log.Printf("Loaded %d users from disk", len(um.users))

	// Ensure default admin exists
	um.ensureAdminLocked()
}

// ensureAdminLocked creates the default admin user if no admin exists.
// Must be called with the lock held.
func (um *UserManager) ensureAdminLocked() {
	// Check if any admin exists
	for _, u := range um.users {
		if u.IsAdmin {
			return // already have an admin
		}
	}
	// Create default admin
	hashedPassword, err := bcrypt.GenerateFromPassword([]byte("admin"), bcrypt.DefaultCost)
	if err != nil {
		log.Printf("Failed to hash admin password: %v", err)
		return
	}
	um.users["admin"] = &User{
		Username:         "admin",
		PasswordHash:     string(hashedPassword),
		Prefs:            Preferences{VisibleRows: 15, VisibleCols: 7},
		IsAdmin:          true,
		CanCreateProject: true,
	}
	um.saveUsersLocked()
	log.Printf("Default admin user created")
}

func (um *UserManager) saveUsersLocked() {
	file, err := os.Create(userPersistenceFile)
	if err != nil {
		log.Printf("Error saving users: %v", err)
		return
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	if err := encoder.Encode(um.users); err != nil {
		log.Printf("Error encoding users: %v", err)
	}
}

// GetPreferences returns the stored preferences for a user
func (um *UserManager) GetPreferences(username string) (Preferences, error) {
	um.mu.RLock()
	defer um.mu.RUnlock()
	u, ok := um.users[username]
	if !ok {
		return Preferences{}, errors.New("user not found")
	}
	return u.Prefs, nil
}

// UpdatePreferences updates and persists the preferences for a user
func (um *UserManager) UpdatePreferences(username string, prefs Preferences) error {
	um.mu.Lock()
	defer um.mu.Unlock()
	u, ok := um.users[username]
	if !ok {
		return errors.New("user not found")
	}
	// sanitize values (reasonable bounds)
	if prefs.VisibleRows <= 0 {
		prefs.VisibleRows = 15
	}
	if prefs.VisibleCols <= 0 {
		prefs.VisibleCols = 7
	}
	u.Prefs = prefs
	um.saveUsersLocked()
	return nil
}

// ChangePassword updates the user's password after verifying the old password
func (um *UserManager) ChangePassword(username, oldPassword, newPassword string) error {
	um.mu.Lock()
	defer um.mu.Unlock()
	u, ok := um.users[username]
	if !ok {
		return errors.New("user not found")
	}
	// verify old password
	if err := bcrypt.CompareHashAndPassword([]byte(u.PasswordHash), []byte(oldPassword)); err != nil {
		return errors.New("invalid current password")
	}
	if len(newPassword) < 6 {
		return errors.New("new password must be at least 6 characters")
	}
	hashedPassword, err := bcrypt.GenerateFromPassword([]byte(newPassword), bcrypt.DefaultCost)
	if err != nil {
		return err
	}
	u.PasswordHash = string(hashedPassword)
	um.saveUsersLocked()
	return nil
}
