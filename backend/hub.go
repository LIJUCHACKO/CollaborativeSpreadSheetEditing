package main

import (
	"encoding/json"
	"log"
)

// Message defines the structure of data exchanged via WebSocket
type Message struct {
	Type    string          `json:"type"`
	SheetID string          `json:"sheet_id"`
	Project string          `json:"project,omitempty"`
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
			roomID := sheetKey(client.projectName, client.sheetID)
			if h.rooms[roomID] == nil {
				h.rooms[roomID] = make(map[*Client]bool)
			}
			h.rooms[roomID][client] = true
			log.Printf("Client registered to sheet %s (project %s): %s", client.sheetID, client.projectName, client.userID)

			// Send current state to the new client
			// In a real app, this should be handled safely with a mutex on the sheet
			sheet := globalSheetManager.GetSheetBy(client.sheetID, client.projectName)
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
				// Also send global chat history independent of sheet
				history := globalChatManager.HistoryFor(client.userID)
				chatPayload, _ := json.Marshal(history)
				client.send <- msgToBytes(&Message{Type: "CHAT_HISTORY", SheetID: "", Payload: chatPayload, User: "system"})

			}

		case client := <-h.unregister:
			roomID := sheetKey(client.projectName, client.sheetID)
			if clients, ok := h.rooms[roomID]; ok {
				if _, ok := clients[client]; ok {
					delete(clients, client)
					close(client.send)
					if len(clients) == 0 {
						delete(h.rooms, roomID)
					}
					log.Printf("Client unregistered from sheet %s (project %s)", client.sheetID, client.projectName)
				}
			}
		case message := <-h.broadcast:
			// The message payload should already be processed/updated in the sheet state before broadcasting?
			// Or we handle the "command" here?
			// For now, let's assume the caller handles state update and we just broadcast
			// Determine final message to send (may differ from inbound command)
			toSend := message

			// Helper: deny non-editors for mutating operations
			denyIfNotEditor := func() bool {
				sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
				if sheet == nil {
					return true
				}
				if !sheet.IsEditor(message.User) {
					deniedPayload, _ := json.Marshal(map[string]string{
						"reason": "not-editor",
						"type":   message.Type,
					})
					// Send denial only to the sender
					toSend = &Message{Type: "EDIT_DENIED", SheetID: message.SheetID, Payload: deniedPayload, User: message.User}
					if clients, ok := h.rooms[sheetKey(message.Project, message.SheetID)]; ok {
						for client := range clients {
							if client.userID != message.User {
								continue
							}
							select {
							case client.send <- msgToBytes(toSend):
							default:
								close(client.send)
								delete(clients, client)
							}
						}
					}
					return true
				}
				return false
			}

			// Persist changes if it's an update
			if message.Type == "UPDATE_CELL" {
				if denyIfNotEditor() {
					continue
				}
				// Parse payload
				var update struct {
					Row   string `json:"row"`
					Col   string `json:"col"`
					Value string `json:"value"`
					User  string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &update); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						sheet.SetCell(update.Row, update.Col, update.Value, message.User)
						// Broadcast updated sheet snapshot
						sheet.mu.RLock()
						payload, _ := json.Marshal(sheet)
						sheet.mu.RUnlock()
						toSend = &Message{
							Type:    "ROW_COL_UPDATED",
							SheetID: message.SheetID,
							Payload: payload,
							User:    message.User,
						}
					}
				} else {
					log.Printf("Error unmarshalling update payload: %v", err)
				}
			} else if message.Type == "UPDATE_CELL_STYLE" {
				if denyIfNotEditor() {
					continue
				}
				var st struct {
					Row        string `json:"row"`
					Col        string `json:"col"`
					Background string `json:"background"`
					Bold       bool   `json:"bold"`
					Italic     bool   `json:"italic"`
					User       string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &st); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						sheet.SetCellStyle(st.Row, st.Col, st.Background, st.Bold, st.Italic, message.User)
						// Broadcast updated sheet snapshot
						sheet.mu.RLock()
						payload, _ := json.Marshal(sheet)
						sheet.mu.RUnlock()
						toSend = &Message{
							Type:    "ROW_COL_UPDATED",
							SheetID: message.SheetID,
							Payload: payload,
							User:    message.User,
						}
					}
				} else {
					log.Printf("Error unmarshalling UPDATE_CELL_STYLE payload: %v", err)
				}
			} else if message.Type == "LOCK_CELL" {
				var req struct {
					Row  string `json:"row"`
					Col  string `json:"col"`
					User string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &req); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						if sheet.LockCell(req.Row, req.Col, message.User) {
							// Broadcast full sheet state
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "ROW_COL_UPDATED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						} else {
							deniedPayload, _ := json.Marshal(map[string]string{
								"row":    req.Row,
								"col":    req.Col,
								"reason": "owner-only",
							})
							toSend = &Message{
								Type:    "LOCK_DENIED",
								SheetID: message.SheetID,
								Payload: deniedPayload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling LOCK_CELL payload: %v", err)
				}
			} else if message.Type == "UNLOCK_CELL" {
				var req struct {
					Row  string `json:"row"`
					Col  string `json:"col"`
					User string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &req); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						if sheet.UnlockCell(req.Row, req.Col, message.User) {
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "ROW_COL_UPDATED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						} else {
							deniedPayload, _ := json.Marshal(map[string]string{
								"row":    req.Row,
								"col":    req.Col,
								"reason": "owner-only",
							})
							toSend = &Message{
								Type:    "UNLOCK_DENIED",
								SheetID: message.SheetID,
								Payload: deniedPayload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling UNLOCK_CELL payload: %v", err)
				}
			} else if message.Type == "RESIZE_COL" {
				if denyIfNotEditor() {
					continue
				}
				var update struct {
					Col   string `json:"col"`
					Width int    `json:"width"`
					User  string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &update); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						sheet.SetColWidth(update.Col, update.Width, message.User)
						// Broadcast updated sheet snapshot
						sheet.mu.RLock()
						payload, _ := json.Marshal(sheet)
						sheet.mu.RUnlock()
						toSend = &Message{
							Type:    "ROW_COL_UPDATED",
							SheetID: message.SheetID,
							Payload: payload,
							User:    message.User,
						}
					}
				} else {
					log.Printf("Error unmarshalling resize col payload: %v", err)
				}
			} else if message.Type == "RESIZE_ROW" {
				if denyIfNotEditor() {
					continue
				}
				var update struct {
					Row    string `json:"row"`
					Height int    `json:"height"`
					User   string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &update); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						sheet.SetRowHeight(update.Row, update.Height, message.User)
						// Broadcast updated sheet snapshot
						sheet.mu.RLock()
						payload, _ := json.Marshal(sheet)
						sheet.mu.RUnlock()
						toSend = &Message{
							Type:    "ROW_COL_UPDATED",
							SheetID: message.SheetID,
							Payload: payload,
							User:    message.User,
						}
					}
				} else {
					log.Printf("Error unmarshalling resize row payload: %v", err)
				}
			} else if message.Type == "MOVE_ROW" {
				if denyIfNotEditor() {
					continue
				}
				var mv struct {
					FromRow   string `json:"fromRow"`
					TargetRow string `json:"targetRow"`
					User      string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &mv); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						moved := sheet.MoveRowBelow(mv.FromRow, mv.TargetRow, message.User)
						if moved {
							// Broadcast updated sheet snapshot
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "ROW_MOVED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling MOVE_ROW payload: %v", err)
				}
			} else if message.Type == "MOVE_COL" {
				if denyIfNotEditor() {
					continue
				}
				var mv struct {
					FromCol   string `json:"fromCol"`
					TargetCol string `json:"targetCol"`
					User      string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &mv); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						moved := sheet.MoveColumnRight(mv.FromCol, mv.TargetCol, message.User)
						if moved {
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "COL_MOVED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling MOVE_COL payload: %v", err)
				}
			} else if message.Type == "INSERT_ROW" {
				if denyIfNotEditor() {
					continue
				}
				var ins struct {
					TargetRow string `json:"targetRow"`
					User      string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &ins); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						inserted := sheet.InsertRowBelow(ins.TargetRow, message.User)
						if inserted {
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "ROW_COL_UPDATED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling INSERT_ROW payload: %v", err)
				}
			} else if message.Type == "INSERT_COL" {
				if denyIfNotEditor() {
					continue
				}
				var ins struct {
					TargetCol string `json:"targetCol"`
					User      string `json:"user"`
				}
				if err := json.Unmarshal(message.Payload, &ins); err == nil {
					sheet := globalSheetManager.GetSheetBy(message.SheetID, message.Project)
					if sheet != nil {
						inserted := sheet.InsertColumnRight(ins.TargetCol, message.User)
						if inserted {
							sheet.mu.RLock()
							payload, _ := json.Marshal(sheet)
							sheet.mu.RUnlock()
							toSend = &Message{
								Type:    "ROW_COL_UPDATED",
								SheetID: message.SheetID,
								Payload: payload,
								User:    message.User,
							}
						}
					}
				} else {
					log.Printf("Error unmarshalling INSERT_COL payload: %v", err)
				}
			} else if message.Type == "SELECTION_COPIED" {
				// Forward selection range/values only to the same user's clients within the sheet room
				// Payload is forwarded as-is; clients will interpret it and render boundaries / clipboard
				toSend = &Message{
					Type:    "SELECTION_SHARED",
					SheetID: message.SheetID,
					Payload: message.Payload,
					User:    message.User,
				}
				//fmt.Println("received selection to broadcast:", string(message.Payload))
				for _, clients := range h.rooms {
					for client := range clients {
						if client.userID != message.User {
							continue
						}
						select {
						case client.send <- msgToBytes(toSend):
						default:
							close(client.send)
							delete(clients, client)
						}
					}
				}
				// Skip the general broadcast below because we already filtered by user
				continue
			} else if message.Type == "CHAT_MESSAGE" {
				// Payload: { text: string, user: string }
				var chat struct {
					Text string `json:"text"`
					User string `json:"user"`
					To   string `json:"to"`
				}
				if err := json.Unmarshal(message.Payload, &chat); err == nil {
					appended := globalChatManager.Append(message.User, chat.Text, chat.To)
					payload, _ := json.Marshal(appended)
					toSend = &Message{Type: "CHAT_APPENDED", SheetID: "", Payload: payload, User: message.User}
					// Broadcast
					if appended.To == "" || appended.To == "all" {
						// broadcast to everyone
						for _, clients := range h.rooms {
							for client := range clients {
								select {
								case client.send <- msgToBytes(toSend):
								default:
									close(client.send)
									delete(clients, client)
								}
							}
						}
					} else {
						// direct message: send to sender and recipient only
						for _, clients := range h.rooms {
							for client := range clients {
								if client.userID != appended.User && client.userID != appended.To {
									continue
								}
								select {
								case client.send <- msgToBytes(toSend):
								default:
									close(client.send)
									delete(clients, client)
								}
							}
						}
					}
					continue // skip per-room broadcast
				} else {
					log.Printf("Error unmarshalling CHAT_MESSAGE payload: %v", err)
				}
			} else if message.Type == "PING" {
				// Optional: reply with a PONG only to sender to confirm connectivity
				toSend = &Message{
					Type:    "PONG",
					SheetID: message.SheetID,
					Payload: nil,
					User:    message.User,
				}
			}

			if clients, ok := h.rooms[sheetKey(message.Project, message.SheetID)]; ok {
				for client := range clients {
					// Don't send back to sender? Or do? usually do for confirmation,
					// but for optimizing latency we might optimistically update frontend.
					// Google docs sends back.
					select {
					case client.send <- msgToBytes(toSend):
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
