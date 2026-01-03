package main

import (
	"encoding/json"
	"log"
)

// Message defines the structure of data exchanged via WebSocket
type Message struct {
	Type    string          `json:"type"`
	SheetID string          `json:"sheet_id"`
	Payload json.RawMessage `json:"payload"`
	User    string          `json:"user,omitempty"` // Username of the sender
}

// Hub maintains the set of active clients and broadcasts messages to the clients.
// Hub maintains the set of active clients and broadcasts messages to the clients.
type Hub struct {
	// Registered clients per sheet.
	rooms map[string]map[*Client]bool

	// Inbound messages from the clients.
	broadcast chan *Message

	// Register requests from the clients.
	register chan *Client

	// Unregister requests from clients.
	unregister chan *Client
}

func newHub() *Hub {
	return &Hub{
		broadcast:  make(chan *Message),
		register:   make(chan *Client),
		unregister: make(chan *Client),
		rooms:      make(map[string]map[*Client]bool),
	}
}

func (h *Hub) run() {
	for {
		select {
		case client := <-h.register:
			if h.rooms[client.sheetID] == nil {
				h.rooms[client.sheetID] = make(map[*Client]bool)
			}
			h.rooms[client.sheetID][client] = true
			log.Printf("Client registered to sheet %s: %s", client.sheetID, client.userID)

			// Send current state to the new client
			// In a real app, this should be handled safely with a mutex on the sheet
			sheet := globalSheetManager.GetSheet(client.sheetID)
			if sheet != nil {
				sheet.mu.RLock()
				payload, _ := json.Marshal(sheet)
				sheet.mu.RUnlock()

				msg := &Message{
					Type:    "INIT",
					Payload: payload,
					User:    "system",
				}
				client.send <- msgToBytes(msg)
			}

		case client := <-h.unregister:
			if clients, ok := h.rooms[client.sheetID]; ok {
				if _, ok := clients[client]; ok {
					delete(clients, client)
					close(client.send)
					if len(clients) == 0 {
						delete(h.rooms, client.sheetID)
					}
					log.Printf("Client unregistered from sheet %s", client.sheetID)
				}
			}
		case message := <-h.broadcast:
			// The message payload should already be processed/updated in the sheet state before broadcasting?
			// Or we handle the "command" here?
			// For now, let's assume the caller handles state update and we just broadcast

			// Persist changes if it's an update
			if message.Type == "UPDATE_CELL" {
				// Parse payload
				// Payload is json.RawMessage ([]byte)
				var update struct {
					Row   string `json:"row"`
					Col   string `json:"col"`
					Value string `json:"value"`
					User  string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &update); err == nil {
					sheet := globalSheetManager.GetSheet(message.SheetID)
					if sheet != nil {
						sheet.SetCell(update.Row, update.Col, update.Value, message.User)
					}
				} else {
					log.Printf("Error unmarshalling update payload: %v", err)
				}
			}

			if clients, ok := h.rooms[message.SheetID]; ok {
				for client := range clients {
					// Don't send back to sender? Or do? usually do for confirmation,
					// but for optimizing latency we might optimistically update frontend.
					// Google docs sends back.
					select {
					case client.send <- msgToBytes(message):
					default:
						close(client.send)
						delete(clients, client)
					}
				}
			}
		}
	}
}

func msgToBytes(msg *Message) []byte {
	b, _ := json.Marshal(msg)
	return b
}
