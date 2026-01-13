package main

import (
	"crypto/rand"
	"encoding/base64"
	"encoding/json"
	"errors"
	"log"
	"os"
	"sync"
	"time"

	"golang.org/x/crypto/bcrypt"
)

const userPersistenceFile = "users.json"
const sessionTimeout = 1 * time.Hour

type User struct {
	Username     string `json:"username"`
	PasswordHash string `json:"password_hash"`
}

type Session struct {
	Token     string
	Username  string
	ExpiresAt time.Time
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

	if _, exists := um.users[username]; exists {
		return errors.New("user already exists")
	}

	hashedPassword, err := bcrypt.GenerateFromPassword([]byte(password), bcrypt.DefaultCost)
	if err != nil {
		return err
	}

	user := &User{
		Username:     username,
		PasswordHash: string(hashedPassword),
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

func (um *UserManager) Load() {
	um.mu.Lock()
	defer um.mu.Unlock()

	file, err := os.Open(userPersistenceFile)
	if err != nil {
		if os.IsNotExist(err) {
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
