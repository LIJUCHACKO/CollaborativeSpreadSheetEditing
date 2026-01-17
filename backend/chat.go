package main

import (
	"encoding/json"
	"log"
	"os"
	"path/filepath"
	"sync"
	"time"
)

type ChatMessage struct {
	Timestamp time.Time `json:"timestamp"`
	User      string    `json:"user"`
	Text      string    `json:"text"`
	To        string    `json:"to,omitempty"` // "all" or specific username
}

type ChatManager struct {
	mu       sync.RWMutex
	messages []ChatMessage
}

var globalChatManager = &ChatManager{}

func chatFilePath() string {
	return filepath.Join(dataDir, "chat.json")
}

func (cm *ChatManager) Load() {
	cm.mu.Lock()
	defer cm.mu.Unlock()
	if err := ensureDataDir(); err != nil {
		log.Printf("chat: ensure data dir: %v", err)
		return
	}
	f, err := os.Open(chatFilePath())
	if err != nil {
		if os.IsNotExist(err) {
			cm.messages = []ChatMessage{}
			return
		}
		log.Printf("chat: open file: %v", err)
		return
	}
	defer f.Close()
	dec := json.NewDecoder(f)
	var msgs []ChatMessage
	if err := dec.Decode(&msgs); err != nil {
		log.Printf("chat: decode: %v", err)
		cm.messages = []ChatMessage{}
		return
	}
	cm.messages = msgs
}

func (cm *ChatManager) Save() {
	cm.mu.RLock()
	defer cm.mu.RUnlock()
	if err := ensureDataDir(); err != nil {
		log.Printf("chat: ensure data dir: %v", err)
		return
	}
	f, err := os.Create(chatFilePath())
	if err != nil {
		log.Printf("chat: create file: %v", err)
		return
	}
	defer f.Close()
	enc := json.NewEncoder(f)
	enc.SetIndent("", "  ")
	if err := enc.Encode(cm.messages); err != nil {
		log.Printf("chat: encode: %v", err)
	}
}

func (cm *ChatManager) Append(user, text, to string) ChatMessage {
	cm.mu.Lock()
	if to == "" {
		to = "all"
	}
	msg := ChatMessage{Timestamp: time.Now(), User: user, Text: text, To: to}
	cm.messages = append(cm.messages, msg)
	cm.mu.Unlock()
	go cm.Save()
	return msg
}

func (cm *ChatManager) History() []ChatMessage {
	cm.mu.RLock()
	defer cm.mu.RUnlock()
	out := make([]ChatMessage, len(cm.messages))
	copy(out, cm.messages)
	return out
}

// HistoryFor returns messages visible to a specific user:
// - broadcast messages (to=="all" or empty)
// - messages sent by the user
// - messages sent to the user
func (cm *ChatManager) HistoryFor(user string) []ChatMessage {
	cm.mu.RLock()
	defer cm.mu.RUnlock()
	out := make([]ChatMessage, 0, len(cm.messages))
	for _, m := range cm.messages {
		to := m.To
		if to == "" || to == "all" || m.User == user || to == user {
			out = append(out, m)
		}
	}
	return out
}
