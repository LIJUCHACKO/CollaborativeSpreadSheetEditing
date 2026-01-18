import React, { useEffect, useState, useRef, useMemo } from 'react';
import { useParams, useNavigate, useLocation } from 'react-router-dom';
import {
    FileSpreadsheet,
    ArrowLeft,
    ArrowUp,
    ArrowDown,
    Save,
    Share2,
    Users,
    History,
    Wifi,
    WifiOff,
    MessageSquare,
    Download,
    Settings,
    Filter
} from 'lucide-react';
import { isSessionValid, clearAuth, getUsername, authenticatedFetch } from '../utils/auth';
import 'bootstrap/dist/css/bootstrap.min.css';

const ROWS = 600;
const COLS = 100;
function toExcelCol(n) {
    let label = '';
    let num = n;
    while (num > 0) {
        num--;
        label = String.fromCharCode(65 + (num % 26)) + label;
        num = Math.floor(num / 26);
    }
    return label;
}
const COL_HEADERS = Array.from({ length: COLS }, (_, i) => toExcelCol(i + 1));
const ROW_HEADERS = Array.from({ length: ROWS }, (_, i) => i + 1);

export default function Sheet() {
    const { id } = useParams();
    const navigate = useNavigate();
    const location = useLocation();
    const username = getUsername();

    // Grid State: map of "row-col" -> Cell
    const [data, setData] = useState({});
    // Audit Log
    const [auditLog, setAuditLog] = useState([]);
    // Connection Status
    const [connected, setConnected] = useState(false);
    const [isSidebarOpen, setSidebarOpen] = useState(false);
    // Add sheetName state
    const [sheetName, setSheetName] = useState('');
    // Project name for this sheet (used for back navigation)
    const [projectName, setProjectName] = useState(() => {
        try {
            const params = new URLSearchParams(location.search);
            return params.get('project') || '';
        } catch {
            return '';
        }
    });
    // Column filters state
    const [filters, setFilters] = useState({});
    const [showFilters, setShowFilters] = useState(false);
    const [focusedCell, setFocusedCell] = useState({ row: 1, col: COL_HEADERS[0] });
    const [isEditing, setIsEditing] = useState(false);
        // Sheet owner state
        const [owner, setOwner] = useState('');
        const [editors, setEditors] = useState([]);
        const isOwner = owner && username && owner === username;
        const canEdit = isOwner || (username && editors.includes(username));
    // Sort configuration: { col: 'A'|'B'|..., direction: 'asc'|'desc' }
    const [sortConfig, setSortConfig] = useState({ col: null, direction: null });
    // Row cut/paste state
    const [cutRow, setCutRow] = useState(null);
    // Column cut/paste state
    const [cutCol, setCutCol] = useState(null);

    // Cell style controls
    const [styleBg, setStyleBg] = useState('');
    const [styleBold, setStyleBold] = useState(false);
    const [styleItalic, setStyleItalic] = useState(false);

    // Multi-cell selection and clipboard state
    const [selectionStart, setSelectionStart] = useState(null); // { row, col }
    const [selectedRange, setSelectedRange] = useState(null); // { startRow, startCol, endRow, endCol }
    const [isSelecting, setIsSelecting] = useState(false);
    const [copiedBlock, setCopiedBlock] = useState(null); // { rows, cols, values: string[][] }
    const [contextMenu, setContextMenu] = useState({ visible: false, x: 0, y: 0, cell: null });
    // Shared selection from other instances of the same user
    const [sharedSelection, setSharedSelection] = useState(null); // { startRow, startCol, endRow, endCol }
    // Chat state
    const [chatMessages, setChatMessages] = useState([]); // [{timestamp, user, text, to}]
    const [chatInput, setChatInput] = useState('');
    const [chatRecipient, setChatRecipient] = useState('all');
    const [allUsers, setAllUsers] = useState([]);
    // Highlight selected audit log entry
    const [selectedAuditId, setSelectedAuditId] = useState(null);
    // Preserve audit log scroll position across open/close
    const auditLogRef = useRef(null);
    const auditLogScrollTopRef = useRef(0);

    const colIndexMap = useMemo(() => {
        const map = {};
        COL_HEADERS.forEach((c, i) => { map[c] = i; });
        return map;
    }, []);

    const colLabelAt = (index) => COL_HEADERS[index] || null;
    const isCellSelected = (rowLabel, colLabel) => {
        if (!selectedRange) return false;
        const rMin = Math.min(selectedRange.startRow, selectedRange.endRow);
        const rMax = Math.max(selectedRange.startRow, selectedRange.endRow);
        const cStartIdx = colIndexMap[selectedRange.startCol] ?? -1;
        const cEndIdx = colIndexMap[selectedRange.endCol] ?? -1;
        const cMin = Math.min(cStartIdx, cEndIdx);
        const cMax = Math.max(cStartIdx, cEndIdx);
        const cIdx = colIndexMap[colLabel] ?? -1;
        return rowLabel >= rMin && rowLabel <= rMax && cIdx >= cMin && cIdx <= cMax;
    };

    const startSelection = (rowLabel, colLabel) => {
        //console.log(rowLabel, colLabel);
        setIsSelecting(true);
        setSelectionStart({ row: rowLabel, col: colLabel });
        setSelectedRange({ startRow: rowLabel, startCol: colLabel, endRow: rowLabel, endCol: colLabel });
        setIsEditing(false);
        setCutRow(null);
        setCutCol(null);
        setContextMenu(prev => ({ ...prev, visible: false }));
    };

    const extendSelection = (rowLabel, colLabel) => {
        if (!isSelecting || !selectionStart) return;
        setSelectedRange(prev => ({ ...prev, endRow: rowLabel, endCol: colLabel }));
    };

    const endSelection = () => {
        if(!connected) return
        if (!isSelecting) return;
        setIsSelecting(false);
        if (!selectedRange) return;

        // Build copied block values from current selected range
        const rMin = Math.min(selectedRange.startRow, selectedRange.endRow);
        const rMax = Math.max(selectedRange.startRow, selectedRange.endRow);
        const cStartIdx = colIndexMap[selectedRange.startCol] ?? -1;
        const cEndIdx = colIndexMap[selectedRange.endCol] ?? -1;
        const cMin = Math.min(cStartIdx, cEndIdx);
        const cMax = Math.max(cStartIdx, cEndIdx);
        const values = [];
        for (let r = rMin; r <= rMax; r++) {
            const rowArr = [];
            for (let ci = cMin; ci <= cMax; ci++) {
                const colLabel = colLabelAt(ci);
                const key = `${r}-${colLabel}`;
                rowArr.push((data[key]?.value ?? '').toString());
            }
            values.push(rowArr);
        }
        // Send selection range and values to backend so other instances of the same user can paste
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = {
                startRow: rMin,
                startCol: colLabelAt(cMin),
                endRow: rMax,
                endCol: colLabelAt(cMax),
                sheet_id: id,
                values,
            };
            ws.current.send(JSON.stringify({ type: 'SELECTION_COPIED', sheet_id: id, payload }));
            //console.log("Send selection:",values);
        }
    };

    const closeContextMenu = () => setContextMenu({ visible: false, x: 0, y: 0, cell: null });

    const showContextMenu = (e, rowLabel, colLabel) => {
        e.preventDefault();
        setIsEditing(false);
        setContextMenu({ visible: true, x: e.clientX, y: e.clientY, cell: { row: rowLabel, col: colLabel } });
    };

    const handleCopySelection = () => {
        if (!selectedRange) return;
        const rMin = Math.min(selectedRange.startRow, selectedRange.endRow);
        const rMax = Math.max(selectedRange.startRow, selectedRange.endRow);
        const cStartIdx = colIndexMap[selectedRange.startCol] ?? -1;
        const cEndIdx = colIndexMap[selectedRange.endCol] ?? -1;
        const cMin = Math.min(cStartIdx, cEndIdx);
        const cMax = Math.max(cStartIdx, cEndIdx);
        const values = [];
        for (let r = rMin; r <= rMax; r++) {
            const rowArr = [];
            for (let ci = cMin; ci <= cMax; ci++) {
                const colLabel = colLabelAt(ci);
                const key = `${r}-${colLabel}`;
                rowArr.push((data[key]?.value ?? '').toString());
            }
            values.push(rowArr);
        }
        //setCopiedBlock({ rows: rMax - rMin + 1, cols: cMax - cMin + 1, values });
        closeContextMenu();
    };

    const handlePasteAt = (anchorRow, anchorColLabel) => {
        if (!copiedBlock || !anchorColLabel) return;
        const anchorIdx = colIndexMap[anchorColLabel] ?? -1;
        if (anchorIdx < 0) return;
        const updates = {};
        let hasConflict = false;
        // Prevent pasting into any locked destination cells
        for (let rOff = 0; rOff < copiedBlock.rows; rOff++) {
            const r = anchorRow + rOff;
            if (r < 1 || r > ROWS) continue;
            for (let cOff = 0; cOff < copiedBlock.cols; cOff++) {
                const cIdx = anchorIdx + cOff;
                if (cIdx < 0 || cIdx >= COLS) continue;
                const cLabel = colLabelAt(cIdx);
                if (!cLabel) continue;
                const key = `${r}-${cLabel}`;
                if (data[key]?.locked) {
                    alert('Cannot paste: destination includes locked cell(s).');
                    closeContextMenu();
                    return;
                }
            }
        }
        for (let rOff = 0; rOff < copiedBlock.rows; rOff++) {
            const r = anchorRow + rOff;
            if (r < 1 || r > ROWS) continue;
            for (let cOff = 0; cOff < copiedBlock.cols; cOff++) {
                const cIdx = anchorIdx + cOff;
                if (cIdx < 0 || cIdx >= COLS) continue;
                const cLabel = colLabelAt(cIdx);
                if (!cLabel) continue;
                const key = `${r}-${cLabel}`;
                const existing = (data[key]?.value ?? '').toString();
                if (existing !== '') hasConflict = true;
            }
        }

        if (hasConflict) {
            const ok = window.confirm('Destination cells contain data. Overwrite?');
            if (!ok) { closeContextMenu(); return; }
        }

        for (let rOff = 0; rOff < copiedBlock.rows; rOff++) {
            const r = anchorRow + rOff;
            if (r < 1 || r > ROWS) continue;
            for (let cOff = 0; cOff < copiedBlock.cols; cOff++) {
                const cIdx = anchorIdx + cOff;
                if (cIdx < 0 || cIdx >= COLS) continue;
                const cLabel = colLabelAt(cIdx);
                if (!cLabel) continue;
                const key = `${r}-${cLabel}`;
                const value = copiedBlock.values[rOff][cOff] ?? '';
                updates[key] = { value, user: username };
            }
        }

        // Apply local state in one batch
        setData(prev => ({ ...prev, ...updates }));

        // Broadcast each cell update to server
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            Object.entries(updates).forEach(([key, cell]) => {
                if (cell.value !== '') {
                    const [rowStr, colLabel] = key.split('-');
                    const payload = { row: rowStr, col: colLabel, value: cell.value, user: username };
                    ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_id: id, payload }));
                }
            });
        }

        closeContextMenu();
    };

    const ws = useRef(null);

    // Viewport state for virtualized grid
    const [cellModified, setCellModified] = useState(0);
    const [rowStart, setRowStart] = useState(1);
    const [visibleRowsCount, setVisibleRowsCount] = useState(15);
    const [colStart, setColStart] = useState(1);
    const [visibleColsCount, setVisibleColsCount] = useState(7);
    const DEFAULT_COL_WIDTH = 112; // px (Tailwind w-28)
    const DEFAULT_ROW_HEIGHT = 40; // px (Tailwind h-10)
    const DEFAULT_ROW_LABEL_WIDTH = 60; // px (Tailwind w-10)
    const DEFAULT_COL_HEADER_HEIGHT = 32; // px
    const [colWidths, setColWidths] = useState(() => {
        const map = {};
        COL_HEADERS.forEach(h => { map[h] = DEFAULT_COL_WIDTH; });
        return map;
    });
    const [rowHeights, setRowHeights] = useState(() => {
        const map = {};
        ROW_HEADERS.forEach(r => { map[r] = DEFAULT_ROW_HEIGHT; });
        return map;
    });
    const [rowLabelWidth, setRowLabelWidth] = useState(DEFAULT_ROW_LABEL_WIDTH);
    const [colHeaderHeight, setColHeaderHeight] = useState(DEFAULT_COL_HEADER_HEIGHT);
    const dragRef = useRef({ type: null, label: null, startPos: 0, startSize: 0 });

    const handleDownloadXlsx = async () => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const projQS = projectName ? `&project=${encodeURIComponent(projectName)}` : '';
            const res = await authenticatedFetch(`http://${host}:8080/api/export?sheet_id=${encodeURIComponent(id)}${projQS}`, {
                method: 'GET',
            });

            if (!res.ok) {
                const text = await res.text();
                alert(`Failed to export sheet: ${text}`);
                return;
            }

            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const safeName = (sheetName || 'sheet') + '.xlsx';
            a.download = safeName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Error downloading XLSX', err);
            alert('An unexpected error occurred while exporting the sheet.');
        }
    };

    const rowEnd = Math.min(rowStart + visibleRowsCount - 1, ROWS);
    const colEnd = Math.min(colStart + visibleColsCount - 1, COLS);

    // Filtered rows state
    const [filteredRowHeaders, setFilteredRowHeaders] = useState(ROW_HEADERS);

    useEffect(() => {
        // Check session validity
        if (!username || !isSessionValid()) {
            clearAuth();
            alert('Your session has expired. Please log in again.');
            navigate('/');
            return;
        }

        // Check session every minute
        const sessionCheckInterval = setInterval(() => {
            if (!isSessionValid()) {
                clearAuth();
                if (ws.current) {
                    ws.current.close();
                }
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        }, 60000); // Check every minute

        // WebSocket connection with reconnection logic
        let reconnectTimeout = null;
        let shouldReconnect = true;

        function connectWS() {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const projQS = projectName ? `&project=${encodeURIComponent(projectName)}` : '';
            const socket = new WebSocket(`ws://${host}:8080/ws?user=${encodeURIComponent(username)}&id=${id}${projQS}` );

            socket.onopen = () => {
                console.log('Connected to WS');
                setConnected(true);

                // SEND initial PING after 5 secs  (connection disconnects in firefox )
                function sendInitialPing() {
                    if (socket.readyState === WebSocket.OPEN) {
                        socket.send(JSON.stringify({ type: 'PING', sheet_id: id }));
                    }
                }

                setTimeout(sendInitialPing, 5000);

            };

            socket.onmessage = (event) => {
                try {
                    //console.log("WS Message:", event.data);

                    // Backend may concatenate multiple JSON messages without a separator.
                    // Split into individual JSON objects and handle each one.
                    const raw = event.data;
                    const messages = [];
                    let depth = 0;
                    let start = -1;
                    for (let i = 0; i < raw.length; i++) {
                        const ch = raw[i];
                        if (ch === '{') {
                            if (depth === 0) start = i;
                            depth++;
                        } else if (ch === '}') {
                            depth--;
                            if (depth === 0 && start !== -1) {
                                messages.push(raw.slice(start, i + 1));
                                start = -1;
                            }
                        }
                    }

                    const parsedMessages = messages.length > 0 ? messages.map(m => JSON.parse(m)) : [JSON.parse(raw)];

                    parsedMessages.forEach((msg) => {
                    if (msg.type === 'INIT') {
                        setInitialState(msg.payload);
                    } else if (msg.type === 'UPDATE_CELL') {
                        const { row, col, value, user } = msg.payload;
                        updateCellState(row, col, value, user);
                    } else if (msg.type === 'RESIZE_COL') {
                        const { col, width } = msg.payload || {};
                        if (col && typeof width === 'number') {
                            setColWidths(prev => ({ ...prev, [col]: width }));
                        }
                    } else if (msg.type === 'RESIZE_ROW') {
                        const { row, height } = msg.payload || {};
                        if (row && typeof height === 'number') {
                            setRowHeights(prev => ({ ...prev, [row]: height }));
                        }
                    } else if (msg.type === 'ROW_MOVED') {
                        setInitialState(msg.payload);
                    } else if (msg.type === 'COL_MOVED') {
                        setInitialState(msg.payload);
                    } else if (msg.type === 'ROW_COL_UPDATED') {
                        setInitialState(msg.payload);
                    } else if (msg.type === 'CHAT_HISTORY') {
                        const list = Array.isArray(msg.payload) ? msg.payload : [];
                        console.log("Chat history:", list);
                        setChatMessages(list);
                    } else if (msg.type === 'CHAT_APPENDED') {
                        const appended = msg.payload;
                        if (appended && appended.text) {
                            setChatMessages(prev => [...prev, appended]);
                        }
                    } else if (msg.type === 'SELECTION_SHARED') {
                        const { startRow, startCol, endRow, endCol, sheet_id,   values } = msg.payload || {};
                        if (startRow && startCol && endRow && endCol) {
                            setSharedSelection({ startRow, startCol, endRow, endCol, sheet_id });
                            if (Array.isArray(values)) {
                                const rows = Math.max(0, endRow - startRow + 1);
                                const cStartIdx = colIndexMap[startCol] ?? -1;
                                const cEndIdx = colIndexMap[endCol] ?? -1;
                                const cols = Math.max(0, Math.abs(cEndIdx - cStartIdx) + 1);
                                setCopiedBlock({ rows, cols, values });
                            }
                        }
                    }else if (msg.type === 'PONG') {
                        console.log("Received PONG from server");
                        setConnected(true);setIsEditing(true);
                    } else if (msg.type === 'EDIT_DENIED') {
                        // Optional UX: show a brief warning when non-editor attempts edit
                        if (!canEdit) {
                            alert('You are not allowed to edit this sheet.');
                        }
                    }
                    });
                } catch (e) {
                    console.error("WS Parse error", e);
                }
            };

            socket.onclose = () => {
                setConnected(false); setIsEditing(false);
                console.log('Disconnected from WS');
                if (shouldReconnect) {
                    // Try to reconnect after 2 seconds
                    reconnectTimeout = setTimeout(() => {
                        connectWS();
                    }, 5000);
                }
            };

            socket.onerror = (e) => {
                console.error('WebSocket error', e);
                // Log current readyState for debugging ping/pong issues
                console.log('WS readyState on error:', socket.readyState);
                // Let onclose handle reconnection; avoid forcing close here unless needed
            };

            ws.current = socket;
        }

        connectWS();

        // Fetch users for chat recipient dropdown
        (async () => {
            try {
                const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
                const res = await authenticatedFetch(`http://${host}:8080/api/users`);
                if (res.ok) {
                    const list = await res.json();
                    if (Array.isArray(list)) setAllUsers(list);
                }
            } catch (e) {
                // ignore fetch errors in chat recipients
            }
        })();

        return () => {
            shouldReconnect = false;
            if (ws.current) ws.current.close();
            if (reconnectTimeout) clearTimeout(reconnectTimeout);
            clearInterval(sessionCheckInterval);
        };
    }, [id, username, navigate]);

    const setInitialState = (sheet) => {
        // Convert nested map to flat key "row-col" if needed
        // sheet.data is map[string]map[string]Cell -> row -> col -> Cell
        const newData = {};
        if (sheet.data) {
            Object.keys(sheet.data).forEach(r => {
                Object.keys(sheet.data[r]).forEach(c => {
                    newData[`${r}-${c}`] = sheet.data[r][c];
                });
            });
        }
        setData(newData);
        if (sheet.audit_log) {
            setAuditLog(sheet.audit_log);
        }
        if (sheet.name) {
            setSheetName(sheet.name);
        } else {
            setSheetName(id);
        }
        if (sheet.project_name) {
            setProjectName(sheet.project_name);
        }
        if (sheet.owner) {
            setOwner(sheet.owner);
        }
        if (sheet.permissions && Array.isArray(sheet.permissions.editors)) {
            setEditors(sheet.permissions.editors);
        }
        // Chat history now comes via CHAT_HISTORY, not from the sheet
        // Apply persisted column widths / row heights if present
        if (sheet.col_widths) {
            setColWidths(prev => ({ ...prev, ...sheet.col_widths }));
        }
        if (sheet.row_heights) {
            setRowHeights(prev => ({ ...prev, ...sheet.row_heights }));
        }
    };

    const updateCellState = (row, col, value, user) => {
        setData(prev => ({
            ...prev,
            [`${row}-${col}`]: { value, user }
        }));
        setCellModified(1)
    };

    const updateCellStyleState = (row, col, background, bold, italic, user) => {
        setData(prev => ({
            ...prev,
            [`${row}-${col}`]: {
                ...(prev[`${row}-${col}`] || {}),
                background,
                bold,
                italic,
                user,
            }
        }));
    };

    const handleCellChange = (r, c, value) => {
        // Optimistic update
        //updateCellState(String(r), String(c), value, username);
        //send update to server only if changed
        if (cellModified === 0) { return; }
        // Send to WB
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const msg = {
                type: 'UPDATE_CELL',
                sheet_id: id,
                payload: { row: String(r), col: String(c), value, user: username }
            };
            ws.current.send(JSON.stringify(msg));
        }
        setCellModified(0);
    };

    const onGlobalMouseMove = (e) => {
        const { type, label, startPos, startSize } = dragRef.current || {};
        if (!type) return;
        if (type === 'col') {
            const delta = e.clientX - startPos;
            const newSize = Math.max(40, startSize + delta);
            setColWidths(prev => ({ ...prev, [label]: newSize }));
            dragRef.current.lastSize = newSize;
        } else if (type === 'row') {
            const delta = e.clientY - startPos;
            const newSize = Math.max(24, startSize + delta);
            setRowHeights(prev => ({ ...prev, [label]: newSize }));
            dragRef.current.lastSize = newSize;
        } else if (type === 'rowLabelWidth') {
            const delta = e.clientX - startPos;
            const newSize = Math.max(30, startSize + delta);
            setRowLabelWidth(newSize);
        } else if (type === 'headerHeight') {
            const delta = e.clientY - startPos;
            const newSize = Math.max(24, startSize + delta);
            setColHeaderHeight(newSize);
        }
    };

    const onGlobalMouseUp = () => {
        const { type, label, lastSize } = dragRef.current || {};
        // Send resize update to server on mouse up
        if ((type === 'col' || type === 'row') && ws.current && ws.current.readyState === WebSocket.OPEN && label && typeof lastSize === 'number' && canEdit) {
            if (type === 'col') {
                const payload = { col: label, width: lastSize, user: username };
                ws.current.send(JSON.stringify({ type: 'RESIZE_COL', sheet_id: id, payload }));
            } else if (type === 'row') {
                const payload = { row: String(label), height: lastSize, user: username };
                ws.current.send(JSON.stringify({ type: 'RESIZE_ROW', sheet_id: id, payload }));
            }
        }
        dragRef.current = { type: null, label: null, startPos: 0, startSize: 0 };
        window.removeEventListener('mousemove', onGlobalMouseMove);
        window.removeEventListener('mouseup', onGlobalMouseUp);
    };

    // Sync toolbar style controls with currently focused cell
    useEffect(() => {
        if (!selectedRange) return;
        //using first cell in selected range if present
        const { row, col } = selectedRange ? { row: selectedRange.startRow, col: selectedRange.startCol } : {};
        if (!row || !col) return;
        const key = `${row}-${col}`;
        const cell = data[key] || {};
        setStyleBg(cell.background || '');
        setStyleBold(!!cell.bold);
        setStyleItalic(!!cell.italic);
    }, [ selectedRange, data]);

    const applyStyleToSelectedRange = () => {
        if (!selectedRange || !canEdit) return;
        const rMin = Math.min(selectedRange.startRow, selectedRange.endRow);
        const rMax = Math.max(selectedRange.startRow, selectedRange.endRow);
        const cStartIdx = colIndexMap[selectedRange.startCol] ?? -1;
        const cEndIdx = colIndexMap[selectedRange.endCol] ?? -1;
        const cMin = Math.min(cStartIdx, cEndIdx);
        const cMax = Math.max(cStartIdx, cEndIdx);

        // Apply locally and broadcast per cell
        for (let r = rMin; r <= rMax; r++) {
            for (let ci = cMin; ci <= cMax; ci++) {
                const colLabel = colLabelAt(ci);
                if (!colLabel) continue;
                const key = `${r}-${colLabel}`;
                const cell = data[key] || {};
                if (cell.locked) continue; // skip locked cells
                updateCellStyleState(r, colLabel, styleBg, styleBold, styleItalic, username);
                if (connected && ws.current && ws.current.readyState === WebSocket.OPEN) {
                    const payload = {
                        row: String(r),
                        col: String(colLabel),
                        background: styleBg || '',
                        bold: !!styleBold,
                        italic: !!styleItalic,
                        user: username,
                    };
                    ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_STYLE', sheet_id: id, payload }));
                }
            }
        }
    };

    const onColResizeMouseDown = (label, e) => {
        e.preventDefault();
        e.stopPropagation();
        dragRef.current = { type: 'col', label, startPos: e.clientX, startSize: colWidths[label] || DEFAULT_COL_WIDTH };
        window.addEventListener('mousemove', onGlobalMouseMove);
        window.addEventListener('mouseup', onGlobalMouseUp);
    };

    const onRowResizeMouseDown = (label, e) => {
        e.preventDefault();
        e.stopPropagation();
        dragRef.current = { type: 'row', label, startPos: e.clientY, startSize: rowHeights[label] || DEFAULT_ROW_HEIGHT };
        window.addEventListener('mousemove', onGlobalMouseMove);
        window.addEventListener('mouseup', onGlobalMouseUp);
    };

    const onRowLabelWidthMouseDown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragRef.current = { type: 'rowLabelWidth', label: 'rowLabel', startPos: e.clientX, startSize: rowLabelWidth };
        window.addEventListener('mousemove', onGlobalMouseMove);
        window.addEventListener('mouseup', onGlobalMouseUp);
    };

    const onHeaderHeightMouseDown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragRef.current = { type: 'headerHeight', label: 'header', startPos: e.clientY, startSize: colHeaderHeight };
        window.addEventListener('mousemove', onGlobalMouseMove);
        window.addEventListener('mouseup', onGlobalMouseUp);
    };

    useEffect(() => {
        return () => {
            window.removeEventListener('mousemove', onGlobalMouseMove);
            window.removeEventListener('mouseup', onGlobalMouseUp);
        };
    }, []);

    useEffect(() => {
        const onWindowMouseUp = () => {};
        const onWindowClick = () => closeContextMenu();
        window.addEventListener('mouseup', onWindowMouseUp);
        window.addEventListener('click', onWindowClick);
        return () => {
            window.removeEventListener('mouseup', onWindowMouseUp);
            window.removeEventListener('click', onWindowClick);
        };
    }, [isSelecting]);

    // When sidebar opens, restore previous scroll position
    useEffect(() => {
        if (isSidebarOpen && auditLogRef.current) {
            auditLogRef.current.scrollTop = auditLogScrollTopRef.current;
        }
    }, [isSidebarOpen]);

    // Update filteredRowHeaders when filters or sort change
    useEffect(() => {
        const activeFilters = Object.entries(filters).filter(([col, val]) => val && val.trim() !== '');
        let newFilteredRowHeaders = activeFilters.length === 0
            ? ROW_HEADERS
            : [
                1,
                ...ROW_HEADERS.filter((rowLabel) => {
                    if (rowLabel === 1) return false; // avoid duplicate, we add 1 explicitly
                    return activeFilters.every(([colLabel, filterVal]) => {
                        const key = `${rowLabel}-${colLabel}`;
                        const cell = data[key] || { value: '' };
                        return String(cell.value).toLowerCase().includes(String(filterVal).toLowerCase());
                    });
                })
            ];

        // Apply sorting if configured
        if (sortConfig && sortConfig.col && sortConfig.direction) {
            const startIdx = newFilteredRowHeaders[0] === 1 ? 1 : 0;
            const rowsToSort = newFilteredRowHeaders.slice(startIdx);

            const parseValue = (row) => {
                const raw = (data[`${row}-${sortConfig.col}`]?.value ?? '').toString().trim();
                const num = parseFloat(raw);
                const isNumeric = raw !== '' && !Number.isNaN(num) && /^-?\d+(?:\.\d+)?$/.test(raw);
                return { raw: raw.toLowerCase(), num, isNumeric };
            };

            rowsToSort.sort((a, b) => {
                const va = parseValue(a);
                const vb = parseValue(b);
                let cmp = 0;
                if (va.isNumeric && vb.isNumeric) {
                    cmp = va.num === vb.num ? 0 : (va.num < vb.num ? -1 : 1);
                } else {
                    cmp = va.raw.localeCompare(vb.raw);
                }
                //if a is empty and b is not, a comes after b
                if (va.raw === '' && vb.raw !== '') {
                    cmp = 1;
                    return cmp;
                } else if (va.raw !== '' && vb.raw === '') {
                    cmp = -1;
                    return cmp;
                }
                

                return sortConfig.direction === 'asc' ? cmp : -cmp;
            });

            newFilteredRowHeaders = startIdx === 1 ? [1, ...rowsToSort] : rowsToSort;
            //console.log("sorted::", newFilteredRowHeaders);
        }

        setFilteredRowHeaders(newFilteredRowHeaders);
    }, [filters, sortConfig, data]);

    // Determine if filtering is active (used to disable row reordering)
    const isFilterActive = Object.values(filters).some(v => v && v.trim() !== '');

    const hasLockedInRow = (rowLabel) => {
        for (let ci = 0; ci < COLS; ci++) {
            const cLabel = colLabelAt(ci);
            const key = `${rowLabel}-${cLabel}`;
            if (data[key]?.locked) return true;
        }
        return false;
    };

    const hasLockedInCol = (colLabel) => {
        for (let r = 1; r <= ROWS; r++) {
            const key = `${r}-${colLabel}`;
            if (data[key]?.locked) return true;
        }
        return false;
    };

    const moveCutRowBelow = (targetRow) => {
        if (cutRow == null) return;
        if (isFilterActive) return; // disabled while filters are active

        // Delegate row move to backend; it will broadcast updated sheet
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { fromRow: String(cutRow), targetRow: String(targetRow), user: username };
            ws.current.send(JSON.stringify({ type: 'MOVE_ROW', sheet_id: id, payload }));
        }

        setCutRow(null);
    };

    const moveCutColRight = (targetCol) => {
        if (cutCol == null) return;
        if (isFilterActive) return; // keep parity with row behavior
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { fromCol: String(cutCol), targetCol: String(targetCol), user: username };
            ws.current.send(JSON.stringify({ type: 'MOVE_COL', sheet_id: id, payload }));
        }
        setCutCol(null);
    };

    const insertRowBelow = (targetRow) => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetRow: String(targetRow), user: username };
            ws.current.send(JSON.stringify({ type: 'INSERT_ROW', sheet_id: id, payload }));
        }
    };

    const insertColumnRight = (targetCol) => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetCol: String(targetCol), user: username };
            ws.current.send(JSON.stringify({ type: 'INSERT_COL', sheet_id: id, payload }));
        }
    };

    // Close sidebar capturing current scroll position
    const closeSidebar = () => {
        if (auditLogRef.current) {
            auditLogScrollTopRef.current = auditLogRef.current.scrollTop;
        }
        setSidebarOpen(false);
    };

    // Toggle sidebar and preserve scroll when closing
    const toggleSidebar = () => {
        if (isSidebarOpen) {
            if (auditLogRef.current) {
                auditLogScrollTopRef.current = auditLogRef.current.scrollTop;
            }
            setSidebarOpen(false);
        } else {
            setSidebarOpen(true);
        }
    };

    // Navigate to a specific cell and ensure it's visible, then focus it
    const navigateToCell = (targetRow, targetColLabel) => {
        if (!targetRow || !targetColLabel) return;
        // Adjust rowStart so targetRow is within visible window
        const rowIdx = filteredRowHeaders.indexOf(targetRow);
        if (rowIdx !== -1) {
            const maxRowStart = Math.max(1, filteredRowHeaders.length - visibleRowsCount + 1);
            const desiredStart = Math.max(1, Math.min(maxRowStart, rowIdx));
            setRowStart(desiredStart);
        }
        // Adjust colStart so targetCol is within visible window
        const colIdx = COL_HEADERS.indexOf(targetColLabel);
        if (colIdx !== -1) {
            const maxColStart = Math.max(1, COLS - visibleColsCount + 1);
            const desiredColStart = Math.max(1, Math.min(maxColStart, colIdx));
            setColStart(desiredColStart);
        }
        // Set focus state and focus the element after re-render
        setFocusedCell({ row: targetRow, col: targetColLabel });
        setIsEditing(false);
        setTimeout(() => {
            const el = document.querySelector(`textarea[data-row="${targetRow}"][data-col="${targetColLabel}"]`);
            if (el) {
                el.focus();
                if (typeof el.scrollIntoView === 'function') el.scrollIntoView({ block: 'center', inline: 'center' });
                if (typeof el.select === 'function') el.select();
            }
        }, 50);
    };

    // Extract row/col from audit log details and navigate
    const navigateToCellFromDetails = (details) => {
        //closeSidebar();
        if (!details || typeof details !== 'string') return;
        // Patterns: "Set cell 28,C to ..." or "Changed cell 4,B from ..."
        const match = details.match(/(?:Set|Changed)\s+cell\s+(\d+),([A-Z]+)\s+/);
        if (match) {
            const row = parseInt(match[1], 10);
            const colLabel = match[2];
            if (!Number.isNaN(row) && COL_HEADERS.includes(colLabel)) {
                navigateToCell(row, colLabel);
            }
            return;
        }
        // Optional: focus column for resize events like "Set width of column C to 93"
        const colMatch = details.match(/column\s+([A-Z]+)\s+/);
        if (colMatch) {
            const colLabel = colMatch[1];
            if (COL_HEADERS.includes(colLabel)) {
                // Focus header row at current first displayed row (use row 1)
                navigateToCell(1, colLabel);
            }
            return;
        }
        // Optional: focus row for resize like "Set height of row 12 to ..."
        const rowMatch = details.match(/row\s+(\d+)\s+/);
        if (rowMatch) {
            const row = parseInt(rowMatch[1], 10);
            if (!Number.isNaN(row)) {
                navigateToCell(row, focusedCell.col);
            }
        }
    };

    const displayedRowHeaders = [
        1,
        ...filteredRowHeaders.slice(
            filteredRowHeaders.length > visibleRowsCount?  rowStart:1,
            Math.min(rowStart + visibleRowsCount , filteredRowHeaders.length )
        )
    ];

    const displayedColHeaders = [COL_HEADERS[0], ...COL_HEADERS.slice(colStart , colEnd)];

    // Clear filter values when showFilters is set to false
    useEffect(() => {
        if (!showFilters) {
            setFilters({});
            sortConfig.direction = null;
            sortConfig.col = null;
            setSortConfig({ ...sortConfig });// to trigger re-render
        }
    }, [showFilters]);

    return (
        <div className="flex h-screen flex-col bg-gray-50 overflow-hidden font-sans text-gray-900">
            {/* Bootstrap Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                    <button
                            onClick={() => {
                                if (projectName) {
                                    navigate(`/project/${encodeURIComponent(projectName)}`);
                                } else {
                                    navigate('/projects');
                                }
                            }}
                            className="btn btn-outline-primary btn-sm d-flex align-items-center"
                        >
                            <ArrowLeft className="me-1" />
                        </button>
                    <span className="navbar-text d-flex align-items-center fw-bold ">
                        <FileSpreadsheet className="me-2" />{sheetName}
                        <span className="mx-3">|</span>
                        <button
                            className="btn btn-outline-primary btn-sm d-flex align-items-center"
                            onClick={handleDownloadXlsx}
                            title="Download as XLSX"
                        >
                            <Download className="me-1" />
                        </button>
                        <button
                            onClick={() => navigate(projectName ? `/settings/${id}?project=${encodeURIComponent(projectName)}` : `/settings/${id}`)}
                            className="btn btn-outline-primary btn-sm d-flex align-items-center ms-2"
                            title="Settings"
                        >
                            <Settings className="me-1" />
                        </button>
                        <button
                            onClick={toggleSidebar}
                            className={`btn btn-outline-primary btn-sm d-flex align-items-center ${isSidebarOpen ? 'active' : ''}`}
                        >
                            <History className="me-1" />
                        </button>
                    </span>
                    <div className="d-flex align-items-center ms-auto">
                        <span className="navbar-text me-4 d-flex align-items-center">
                            <i className="bi bi-person me-1" /> {username}
                        </span>
                        <span className={`navbar-text d-flex align-items-center fw-bold ${connected ? 'text-success' : 'text-danger'}`}
                              title={connected ? 'Connected' : 'Offline'}>
                            {connected ? <Wifi className="me-1" size={18} /> : <WifiOff className="me-1" size={18} />}
                            {connected ? 'Live' : 'Offline'}
                        </span>
                    </div>
                </div>
            </nav>
            {/* Header / Toolbar */}
            <header className="bg-white border-b border-gray-200 shadow-sm z-20">
                <div className="flex items-center justify-between px-4 h-16">
                    <div className="flex items-center gap-2">
                        <button
                            className="px-3 py-1.5 text-sm rounded border border-gray-300 bg-white hover:bg-gray-100 flex items-center gap-2"
                            onClick={() => setShowFilters(v => !v)}
                            title="Toggle column filters"
                        >
                            <Filter size={16} />
                            {showFilters ? 'Hide Filters' : 'Show Filters'}
                        </button>
                        

                            <span className="text-sm text-gray-600">Rows visible</span>
                            <input
                                type="number"
                                className="w-16 border rounded px-2 py-1 text-sm"
                                min={1}
                                max={ROWS}
                                value={visibleRowsCount}
                                onChange={(e) => {
                                    const val = Math.max(1, Math.min(ROWS, parseInt(e.target.value, 10) || 1));
                                    setVisibleRowsCount(val);
                                    setRowStart((prev) => Math.min(prev, Math.max(2, ROWS - val + 1)));
                                }}
                                title="Visible rows"
                            />
                            <span className="text-sm text-gray-600 ml-2">Cols visible</span>
                            <input
                                type="number"
                                className="w-16 border rounded px-2 py-1 text-sm"
                                min={1}
                                max={COLS}
                                value={visibleColsCount}
                                onChange={(e) => {
                                    const val = Math.max(1, Math.min(COLS, parseInt(e.target.value, 10) || 1));
                                    setVisibleColsCount(val);
                                    setColStart((prev) => Math.min(prev, Math.max(1, COLS - val + 1)));
                                }}
                                title="Visible columns"
                            />
                            {/* Style controls for focused cell */}
                        <div className="flex items-center gap-2 ml-2">
                            <span className="text-sm text-gray-600">Bg</span>
                            <input
                                type="color"
                                value={styleBg || '#ffffff'}
                                onChange={(e) => setStyleBg(e.target.value)}
                                disabled={!canEdit}
                                title="Background color"
                            />
                            <button
                                className={`px-2 py-1 text-sm rounded border ${styleBold ? 'bg-indigo-100 border-indigo-300' : 'border-gray-300 bg-white'} hover:bg-gray-100`}
                                onClick={() => setStyleBold(v => !v)}
                                disabled={!canEdit}
                                title="Bold"
                            >
                                B
                            </button>
                            <button
                                className={`px-2 py-1 text-sm rounded border ${styleItalic ? 'bg-indigo-100 border-indigo-300' : 'border-gray-300 bg-white'} hover:bg-gray-100`}
                                onClick={() => setStyleItalic(v => !v)}
                                disabled={!canEdit}
                                title="Italic"
                            >
                                I
                            </button>
                            <button
                                className="px-2 py-1 text-sm rounded border border-gray-300 bg-white hover:bg-gray-100"
                                onClick={applyStyleToSelectedRange}
                                disabled={!canEdit}
                                title="Apply to selected cells"
                            >
                                Apply
                            </button>
                        </div>
                        
                    </div>
                </div>
            </header>

            <div style={{ display: 'inline' ,float: 'left'}} className="flex flex-1 overflow-hidden relative">
                {/* Sidebar / Audit Log */}
                {isSidebarOpen && (
                     <div style={{ position: 'fixed', right: 16, top: 70, width: 360, height: 'calc(70% - 32px)', zIndex: 1100 }}>
                   
                        <div className="d-flex justify-content-between align-items-center p-3 border-bottom bg-light">
                            <h5 className="mb-0 d-flex align-items-center">
                                <History className="me-2" size={18} /> Activity Log
                            </h5>
                            <button 
                                onClick={closeSidebar} 
                                className="btn btn-sm btn-light"
                                aria-label="Close sidebar"
                            >
                                <ArrowLeft size={18} />
                            </button>
                        </div>
                        <div ref={auditLogRef} className="overflow-auto p-3" style={{ height: 'calc(70% - 56px)' ,overflowY:'scroll'}}>
                            {auditLog.slice().reverse().map((entry, i) => {
                                const entryId = `${entry.timestamp || i}|${entry.user || ''}|${entry.action || ''}|${entry.details || ''}`;
                                const isSelected = selectedAuditId === entryId;
                                return (
                                <div key={entryId} className={`d-flex gap-2 mb-3 p-2 rounded hover-bg-light ${isSelected ? 'bg-green-500 border border-green-200' : ''}`} style={{ cursor: 'pointer' }} onClick={() => { setSelectedAuditId(entryId); navigateToCellFromDetails(entry.details || entry.action); }}>
                                    <div className="flex-shrink-0">
                                        <div 
                                            className="rounded-circle bg-gradient d-flex align-items-center justify-content-center text-white fw-bold"
                                            style={{ 
                                                width: '32px', 
                                                height: '32px', 
                                                fontSize: '12px',
                                                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
                                            }}
                                        >
                                            {entry.user?.charAt(0).toUpperCase()}
                                        </div>
                                    </div>
                                    <div className="flex-grow-1">
                                        <div className="d-flex align-items-center gap-2 mb-1">
                                            <span className="fw-semibold small">{entry.user}</span>
                                            <span className="text-muted" style={{ fontSize: '0.75rem' }}>
                                                {new Date(entry.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                                            </span>
                                        </div>
                                        <p className="mb-0 text-muted small">
                                            {entry.details || entry.action}
                                        </p>
                                    </div>
                                </div>
                                );
                            })}
                            {auditLog.length === 0 && (
                                <div className="text-center text-muted py-5">
                                    <History className="mb-2" size={48} opacity={0.3} />
                                    <p className="mb-0">No activity yet.</p>
                                </div>
                            )}
                        </div>
                    </div>
                )}
                {/* Grid Area */}
                <div className="flex-1 overflow-hidden p-6 bg-gray-100/50" >
                    {/* Scrollbars + Grid layout */}
                    <div className="h-full w-full" style={{ display: 'grid', gridTemplateColumns: '24px auto', gridTemplateRows: '24px auto 24px' }}>
                        
                        {/* Top horizontal column scrollbar */}
                        <div style={{ gridColumn: '2 / span 1', gridRow: '1 / span 1' }}
                             onWheel={(e) => {
                                 e.preventDefault();
                                 // Commit edit and exit editing mode on scroll
                                 setIsEditing(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, COLS - visibleColsCount + 1);
                                 setColStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}>
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, COLS - visibleColsCount + 1)}
                                value={colStart}
                                onChange={(e) =>{ 
                                    // Commit edit and exit editing mode on scroll
                                    setIsEditing(false);
                                    const { row, col } = focusedCell;
                                    const key = `${row}-${col}`;
                                    if (data[key]) {
                                        handleCellChange(row, col, data[key].value);
                                    }
                                    setColStart(Math.max(1, Math.min(COLS - visibleColsCount + 1, parseInt(e.target.value, 10) || 1)))}}
                                style={{ width: '100%' }}
                                aria-label="Columns scrollbar"
                            />
                        </div>

                        {/* Left vertical row scrollbar */}
                        <div style={{ gridColumn: '1 / span 1', gridRow: '2 / span 1' }}
                             onWheel={(e) => {
                                 e.preventDefault();
                                 setIsEditing(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, ROWS - visibleRowsCount + 1);
                                 setRowStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}
                             className="flex items-stretch">
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, ROWS - visibleRowsCount + 1)}
                                value={rowStart}
                                onChange={(e) => {
                                    setIsEditing(false);
                                    const { row, col } = focusedCell;
                                    const key = `${row}-${col}`;
                                    if (data[key]) {
                                        handleCellChange(row, col, data[key].value);
                                    }
                                    
                                    setRowStart(Math.max(1, Math.min(ROWS - visibleRowsCount + 1, parseInt(e.target.value, 10) || 1)))}}
                                style={{ writingMode: 'vertical-rl', height: '100%', width: '100%' }}
                                aria-label="Rows scrollbar"
                            />
                        </div>

                        {/* Grid content */}
                        <div style={{ gridColumn: '2 / span 1', gridRow: '2 / span 1', overflow: 'auto' }}
                        onWheel={(e) => {
                                 e.preventDefault();
                                 setIsEditing(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, ROWS - visibleRowsCount + 1);
                                 setRowStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}
                        >
                            <div className="inline-block bg-blue-500 rounded-lg shadow-lg border border-gray-200 overflow-hidden">
                        <table className="border-collapse" >
                            <thead>
                                <tr>
                                    <th
                                        className="bg-gray-50 border-b border-r border-gray-200 p-2 relative select-none"
                                        style={{ width: `${rowLabelWidth}px`, height: `${colHeaderHeight}px` }}
                                    >
                                        
                                    </th>
                                    {displayedColHeaders.map(h => (
                                        <th
                                            key={h}
                                            className="bg-gray-50 border-b border-r border-gray-200 p-2 text-xs font-semibold text-gray-500 uppercase tracking-wider text-center select-none relative"
                                            style={{position: 'relative', width: `${colWidths[h] || DEFAULT_COL_WIDTH}px`, height: `${colHeaderHeight}px` ,padding :`0`}}
                                        >
                                            <div className="flex items-center justify-center gap-1">
                                                <span>{h}</span>
                                                <div style={{ position: 'absolute', top: 2, left: 2, display: 'flex', gap: '4px', zIndex: 25 }}>
                                                    {connected && canEdit && (
                                                        <button
                                                            type="button"
                                                            className="btn btn-xs btn-light"
                                                            disabled={isFilterActive}
                                                            title={isFilterActive ? 'Disabled while filters are active' : `Insert column to the right of ${h}`}
                                                            onClick={() => insertColumnRight(h)}
                                                            style={{ padding: '0 4px', fontSize: '10px' }}
                                                        >
                                                            <span role="img" aria-label="insert-col">➕</span>
                                                        </button>
                                                    )}
                                                    { cutCol == null && connected && canEdit && (
                                                        <button
                                                            type="button"
                                                            className="btn btn-xs btn-light"
                                                            disabled={isFilterActive}
                                                            title={isFilterActive ? 'Disabled while filters are active' : `Cut column ${h}`}
                                                            onClick={() => {
                                                                if (hasLockedInCol(h)) { alert('Cannot cut: column has locked cell(s).'); return; }
                                                                setCutCol(h); setCutRow(null);
                                                            }}
                                                            style={{ padding: '0 4px', fontSize: '10px' }}
                                                        >
                                                            <span role="img" aria-label="cut">✂️</span>
                                                        </button>
                                                    )}
                                                    {cutCol != null && cutCol !== h && connected && canEdit && (
                                                        <button
                                                            type="button"
                                                            className="btn btn-xs btn-light"
                                                            disabled={isFilterActive}
                                                            title={isFilterActive ? 'Disabled while filters are active' : `Insert cut column to the right of ${h}`}
                                                            onClick={() => { moveCutColRight(h); setCutRow(null); setCutCol(null); }}
                                                            style={{ padding: '0 4px', fontSize: '10px' }}
                                                        >
                                                            <span role="img" aria-label="paste">📋</span>
                                                        </button>
                                                    )}
                                                </div>
                                            </div>
                                            <span
                                                onMouseDown={(e) => onColResizeMouseDown(h, e)}
                                                title="Drag to resize column"
                                                role="separator"
                                                aria-orientation="vertical"
                                                style={{
                                                    position: 'absolute',
                                                    top: 0,
                                                    right: 0,
                                                    width: '8px',
                                                    height: '100%',
                                                    cursor: 'col-resize',
                                                    userSelect: 'none',
                                                    background: 'rgba(99,102,241,0.15)', // indigo-500 tint
                                                    borderRight: '1px solid #6366f1',
                                                    zIndex: 20,
                                                    touchAction: 'none'
                                                }}
                                            ></span>
                                        </th>
                                    ))}
                                </tr>
                                {showFilters && (
                                    <tr>
                                        <th
                                            className="bg-gray-50 border-b border-r border-gray-200 p-1 text-xs text-gray-500 text-center select-none"
                                            style={{ width: `${rowLabelWidth}px` }}
                                        >
                                            #
                                        </th>
                                        {displayedColHeaders.map((h) => (
                                            <th
                                                key={`filter-${h}`}
                                                className="bg-gray-50 border-b border-r border-gray-200 p-1 inline-flex items-center gap-1"
                                                style={{ width: `${colWidths[h] || DEFAULT_COL_WIDTH}px` , position: 'relative', padding: '0' }}
                                            >
                                                <input
                                                    type="text"
                                                    className="px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:border-indigo-500"
                                                    placeholder="Filter"
                                                    value={filters[h] || ''}
                                                    onChange={(e) => setFilters(prev => ({ ...prev, [h]: e.target.value }))}
                                                    style={{ width: `${colWidths[h] || DEFAULT_COL_WIDTH}px`, padding: '0' }}
                                                />
                                                <span
                                                    style={{
                                                    position: 'absolute',
                                                    top: 0,
                                                    right: 0
                                                    }}
                                                
                                                >
                                                    <button
                                                        type="button"
                                                        className={`p-0.5 rounded ${sortConfig.col === h && sortConfig.direction === 'asc' ? 'bg-indigo-100 text-indigo-600' : 'text-gray-500 hover:text-indigo-600'}`}
                                                        title="Sort ascending"
                                                        onClick={() => {
                                                            if (sortConfig.col === h && sortConfig.direction === 'asc') {
                                                                setSortConfig({ col: null, direction: null });
                                                            } else {
                                                                setSortConfig({ col: h, direction: 'asc' });
                                                            }
                                                        }}
                                                    >
                                                        <ArrowUp size={12} />
                                                    </button>
                                                    <button
                                                        type="button"
                                                        className={`p-0.5 rounded ${sortConfig.col === h && sortConfig.direction === 'desc' ? 'bg-indigo-100 text-indigo-600' : 'text-gray-500 hover:text-indigo-600'}`}
                                                        title="Sort descending"
                                                        onClick={() => {
                                                            if (sortConfig.col === h && sortConfig.direction === 'desc') {
                                                                setSortConfig({ col: null, direction: null });
                                                            } else {
                                                                setSortConfig({ col: h, direction: 'desc' });
                                                            }
                                                        }}
                                                    >
                                                        <ArrowDown size={12} />
                                                    </button>
                                                </span>
                                            </th>
                                        ))}
                                    </tr>
                                )}
                            </thead>
                            <tbody>
                                {displayedRowHeaders.map((rowLabel) => (
                                    <tr key={rowLabel}>
                                        <td
                                            className="bg-gray-50 border-b border-r border-gray-200 p-2 text-right text-xs font-semibold text-gray-500 select-none relative"
                                            style={{ position: 'relative',height: `${rowHeights[rowLabel] || DEFAULT_ROW_HEIGHT}px`, width: `${rowLabelWidth}px`,padding :`0` }}
                                        >
                                            
                                            {/* Row actions: Insert / Cut / Paste */}
                                            <div style={{ position: 'absolute', top: 0, left: 0, display: 'flex', gap: '4px', zIndex: 25 }}>
                                                {connected && canEdit && (
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Insert row below ${rowLabel}`}
                                                        onClick={() => insertRowBelow(rowLabel)}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="insert-row">➕</span>
                                                    </button>
                                                )}
                                                {cutRow === null && connected && canEdit &&(<button
                                                    type="button"
                                                    className="btn btn-xs btn-light"
                                                    disabled={isFilterActive}
                                                    title={isFilterActive ? 'Disabled while filters are active' : 'Cut this row'}
                                                    onClick={() => {
                                                        if (hasLockedInRow(rowLabel)) { alert('Cannot cut: row has locked cell(s).'); return; }
                                                        setCutRow(rowLabel); setCutCol(null);
                                                    }}
                                                    style={{ padding: '0 0px', fontSize: '8px' }}
                                                >
                                                    <span role="img" aria-label="cut">✂️</span>
                                                </button>)}
                                                {cutRow != null && cutRow !== rowLabel && connected && canEdit &&(
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Insert cut row below row ${rowLabel}`}
                                                        onClick={() => { moveCutRowBelow(rowLabel); setCutRow(null); setCutCol(null); }}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="paste">📋</span>
                                                    </button>
                                                )}
                                                
                                            </div>
                                            <span>{rowLabel}</span>
                                             <div
                                                onMouseDown={(e) => onRowResizeMouseDown(rowLabel, e)}
                                                title="Drag to resize row"
                                                role="separator"
                                                aria-orientation="horizontal"
                                                style={{
                                                    position: 'absolute',
                                                    left: 0,
                                                    bottom: 0,
                                                    width: '100%',
                                                    height: '8px',
                                                    cursor: 'row-resize',
                                                    userSelect: 'none',
                                                    background: 'rgba(99,102,241,0.15)',
                                                    borderTop: '1px solid #0ead23ff',
                                                    zIndex: 20,
                                                    touchAction: 'none'
                                                }}
                                            ></div>
                                        </td>
                                        {displayedColHeaders.map((colLabel) => {
                                            // Only render cell if sheetId matches current id
                                            
                                            const key = `${rowLabel}-${colLabel}`;
                                            const cell = data[key] || { value: '' };
                                            //const selected = isCellSelected(rowLabel, colLabel);
                                            const inShared = sharedSelection ? (function(){
                                                const rMin = Math.min(sharedSelection.startRow, sharedSelection.endRow);
                                                const rMax = Math.max(sharedSelection.startRow, sharedSelection.endRow);
                                                const cStartIdx = colIndexMap[sharedSelection.startCol] ?? -1;
                                                const cEndIdx = colIndexMap[sharedSelection.endCol] ?? -1;
                                                const cMin = Math.min(cStartIdx, cEndIdx);
                                                const cMax = Math.max(cStartIdx, cEndIdx);
                                                const cIdx = colIndexMap[colLabel] ?? -1;
                                                return rowLabel >= rMin && rowLabel <= rMax && cIdx >= cMin && cIdx <= cMax;
                                            })() : false;
                                            const boundaryStyles = (function(){
                                                if (!sharedSelection) return {};
                                                // Disable shared selection borders if it belongs to a different sheet
                                                //console.log('sharedSelection.sheet_id', sharedSelection.sheet_id, 'current id', id);
                                                if (sharedSelection.sheet_id && sharedSelection.sheet_id !== id) return {};
                                                const rMin = Math.min(sharedSelection.startRow, sharedSelection.endRow);
                                                const rMax = Math.max(sharedSelection.startRow, sharedSelection.endRow);
                                                const cStartIdx = colIndexMap[sharedSelection.startCol] ?? -1;
                                                const cEndIdx = colIndexMap[sharedSelection.endCol] ?? -1;
                                                const cMin = Math.min(cStartIdx, cEndIdx);
                                                const cMax = Math.max(cStartIdx, cEndIdx);
                                                const cIdx = colIndexMap[colLabel] ?? -1;
                                                const color = '#6366f1';
                                                const style = {};
                                                if (rowLabel === rMin && cIdx >= cMin && cIdx <= cMax) style.borderTop = `2px solid ${color}`;
                                                if (rowLabel === rMax && cIdx >= cMin && cIdx <= cMax) style.borderBottom = `2px solid ${color}`;
                                                if (cIdx === cMin && rowLabel >= rMin && rowLabel <= rMax) style.borderLeft = `2px solid ${color}`;
                                                if (cIdx === cMax && rowLabel >= rMin && rowLabel <= rMax) style.borderRight = `2px solid ${color}`;
                                                return style;
                                            })();

                                            return (
                                                <td
                                                    key={key}
                                                    className={`border-b border-r bg-gray-100 p-0 relative min-w-[7rem] group ${ inShared ? 'bg-indigo-50' : 'bg-white'} hover:bg-green-50/20 transition-colors`}
                                                    style={{ width: `${colWidths[colLabel] || DEFAULT_COL_WIDTH}px`, height: `${rowHeights[rowLabel] || DEFAULT_ROW_HEIGHT}px`, ...boundaryStyles }}
                                                    onContextMenu={(e) => {  !isEditing && showContextMenu(e, rowLabel, colLabel)}}
                                                >
                                                    <textarea
                                                        className={`w-full h-full px-3 py-1 text-sm outline-none border-2 border-transparent focus:border-green-100 focus:ring-0 z-10 relative  text-gray-800 resize-none`}
                                                        rows={1}
                                                        style={{
                                                            width: '100%',
                                                            height: '100%',
                                                            boxSizing: 'border-box',
                                                            display: 'block',
                                                            overflow: 'auto',
                                                            resize: 'none',
                                                            whiteSpace: 'pre-wrap',
                                                            backgroundColor: (cell.background && cell.background !== '') ? cell.background : undefined,
                                                            fontWeight: cell.bold ? '700' : 'normal',
                                                            fontStyle: cell.italic ? 'italic' : 'normal',
                                                        }}
                                                        value={cell.value}
                                                        data-row={rowLabel}
                                                        data-col={colLabel}
                                                        readOnly={!!cell.locked || !canEdit}
                                                        onFocus={() => { setFocusedCell({ row: rowLabel, col: colLabel }); setIsEditing(false); }}
                                                        onMouseOver={e => { e.target.focus(); }}
                                                        onDoubleClick={(e) => {
                                                            if (isEditing) return;
                                                            if (cell.locked || !canEdit) return;
                                                            // Prevent default double-click text selection
                                                            e.preventDefault();
                                                            e.target.focus();
                                                            if (connected) {
                                                                setIsEditing(true);
                                                                setCutRow(null);
                                                                setCutCol(null);
                                                            }
                                                            // Clear any selection by collapsing caret
                                                            if (typeof e.target.setSelectionRange === 'function') {
                                                                const len = e.target.value.length;
                                                                e.target.setSelectionRange(len, len);
                                                            }
                                                        }}
                                                        onMouseDown={(e) => { 
                                                            if (isEditing) {
                                                                // In edit mode: allow normal text selection, but keep focus
                                                                // Do NOT call preventDefault or selection handlers
                                                                return;
                                                            }
                                                            e.preventDefault();
                                                            e.target.focus();
                                                            if (e.button === 0 ) {
                                                                startSelection(rowLabel, colLabel);
                                                            }
                                                        }}
                                                        onMouseEnter={() => { if(!isEditing) extendSelection(rowLabel, colLabel);}   } 
                                                        onMouseUp={(e) => {
                                                                    if(!isEditing) {
                                                                        extendSelection(rowLabel, colLabel);
                                                                        endSelection(); 
                                                                    }
                                                        }}
                                                      
                                                        onKeyDown={(e) => {
                                                            const keys = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'];
                                                            // Enter edit mode when typing any non-arrow key (including Enter)
                                                            if (!keys.includes(e.key)) { if(cell.locked) return; if(connected) setIsEditing(true); return; }
                                                            // In edit mode, allow default arrow behavior inside textarea and disable cell navigation
                                                            if (isEditing && keys.includes(e.key)) { return; }
                                                            e.preventDefault();
                                                            let nextRow = rowLabel;
                                                            let nextCol = colLabel;
                                                            const rowIdx = filteredRowHeaders.indexOf(rowLabel);
                                                            const colIdx = COL_HEADERS.indexOf(colLabel);
                                                            
                                                            if (e.key === 'ArrowDown') {
                                                                if (rowIdx !== -1 && rowIdx + 1 < filteredRowHeaders.length) {
                                                                    nextRow = filteredRowHeaders[rowIdx + 1];
                                                                    const currentRowEnd = Math.min(rowStart + visibleRowsCount , filteredRowHeaders.length );
                                                                    if (rowIdx + 1 > currentRowEnd - 1) {
                                                                        setRowStart(prev => Math.min(filteredRowHeaders.length - visibleRowsCount + 1 > 1 ?  filteredRowHeaders.length - visibleRowsCount + 1 : 1, rowStart+1));
                                                                    }
                                                                }
                                                            } else if (e.key === 'ArrowUp') {
                                                                if (rowIdx > 0) {
                                                                    nextRow = filteredRowHeaders[rowIdx - 1];
                                                                    if (rowIdx -1  < rowStart + 1 && rowIdx > 0) {
                                                                        setRowStart(prev => Math.max(1, rowIdx - 1));
                                                                    }
                                                                }
                                        
                                                            } else if (e.key === 'ArrowRight') {
                                                                if (colIdx !== -1 && colIdx + 1 < COL_HEADERS.length) {
                                                                    nextCol = COL_HEADERS[colIdx + 1];
                                                                    const nextColNum = colIdx + 2;
                                                                    if (nextColNum > colEnd) {
                                                                        setColStart(prev => Math.min(COLS - visibleColsCount + 1, prev + 1));
                                                                    }
                                                                }
                                                            } else if (e.key === 'ArrowLeft') {
                                                                if (colIdx > 0) {
                                                                    nextCol = COL_HEADERS[colIdx - 1];
                                                                    const nextColNum = colIdx;
                                                                    if (nextColNum <= colStart) {
                                                                        setColStart(prev => Math.max(1, prev - 1));
                                                                    }
                                                                }
                                                            }
                                                            setFocusedCell({ row: nextRow, col: nextCol });
                                                            setTimeout(() => {
                                                                const el = document.querySelector(`textarea[data-row="${nextRow}"][data-col="${nextCol}"]`);
                                                                if (el) {
                                                                    el.focus();
                                                                    if (typeof el.select === 'function') el.select();
                                                                }
                                                            }, 0);
                                                            }}
                                                        // Only update value locally while editing, commit on blur
                                                        onChange={(e) => {
                                                            // Update local state for textarea value
                                                            if (cell.locked || !canEdit) return;
                                                            if (connected)
                                                            updateCellState(rowLabel, colLabel, e.target.value);
                                                            
                                                        }}
                                                        onBlur={(e) => {
                                                            setIsEditing(false);
                                                            // Commit value to backend only on blur
                                                            if (!cell.locked && canEdit) handleCellChange(rowLabel, colLabel, e.target.value);
                                                        }}
                                                    />

                                                    
                                                        {/* Context Menu */}
                                                        {contextMenu.visible && (
                                                            <div
                                                                style={{ position: 'fixed', top: contextMenu.y, left: contextMenu.x, zIndex: 2000 }}
                                                                className="bg-white border shadow rounded-md p-2 text-sm"
                                                                onClick={(e) => e.stopPropagation()}
                                                            >
                                                                <button
                                                                    className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                    disabled={!selectedRange}
                                                                    onClick={handleCopySelection}
                                                                >
                                                                    Copy
                                                                </button>
                                                                <button
                                                                    className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                    disabled={!copiedBlock || !contextMenu.cell}
                                                                    onClick={() => handlePasteAt(contextMenu.cell.row, contextMenu.cell.col)}
                                                                >
                                                                    Paste
                                                                </button>
                                                                {isOwner && contextMenu.cell && (
                                                                    <>
                                                                        {!data[`${contextMenu.cell.row}-${contextMenu.cell.col}`]?.locked && (
                                                                            <button
                                                                                className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                onClick={() => {
                                                                                    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                                                                                        const payload = { row: String(contextMenu.cell.row), col: String(contextMenu.cell.col), user: username };
                                                                                        ws.current.send(JSON.stringify({ type: 'LOCK_CELL', sheet_id: id, payload }));
                                                                                    }
                                                                                    closeContextMenu();
                                                                                }}
                                                                            >
                                                                                Lock Cell
                                                                            </button>
                                                                        )}
                                                                        {data[`${contextMenu.cell.row}-${contextMenu.cell.col}`]?.locked && (
                                                                            <button
                                                                                className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                onClick={() => {
                                                                                    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                                                                                        const payload = { row: String(contextMenu.cell.row), col: String(contextMenu.cell.col), user: username };
                                                                                        ws.current.send(JSON.stringify({ type: 'UNLOCK_CELL', sheet_id: id, payload }));
                                                                                    }
                                                                                    closeContextMenu();
                                                                                }}
                                                                            >
                                                                                Unlock Cell
                                                                            </button>
                                                                        )}
                                                                    </>
                                                                )}
                                                            </div>
                                                        )}
                                                    {cell.user && cell.user !== username && (
                                                        <div className="absolute top-0 right-0 w-0 h-0 border-t-[8px] border-r-[8px] border-t-purple-500 border-r-transparent transform rotate-90" title={`Edited by ${cell.user}`}></div>
                                                    )}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                            </div>
                        </div>

                        {/* Bottom horizontal column scrollbar (alias) */}
                        <div style={{ gridColumn: '2 / span 1', gridRow: '3 / span 1' }}
                             onWheel={(e) => {
                                 e.preventDefault();
                                 // Commit edit and exit editing mode on scroll
                                 setIsEditing(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, COLS - visibleColsCount + 1);
                                 setColStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}>
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, COLS - visibleColsCount + 1)}
                                value={colStart}
                                onChange={(e) =>{
                                    // Commit edit and exit editing mode on scroll
                                    setIsEditing(false);
                                    const { row, col } = focusedCell;
                                    const key = `${row}-${col}`;
                                    if (data[key]) {
                                        handleCellChange(row, col, data[key].value);
                                    }
                                    setColStart(Math.max(1, Math.min(COLS - visibleColsCount + 1, parseInt(e.target.value, 10) || 1)))}}
                                style={{ width: '100%' }}
                                aria-label="Columns scrollbar (bottom)"
                            />
                        </div>

                        {/* Chat panel (fixed bottom-right) */}
                        <div style={{ position: 'fixed', right: 16, bottom: 16, width: 360, zIndex: 1100 }}>
                            <div className="card shadow-sm">
                                <div className="card-header py-2 d-flex align-items-center justify-content-between">
                                    <span className="fw-semibold small d-flex align-items-center"><MessageSquare size={16} className="me-2"/> Chat</span>
                                    <span className="badge bg-light text-dark">{chatMessages.length}</span>
                                </div>
                                <div className="card-body p-2" style={{ maxHeight: 240, overflowY: 'auto' }}>
                                    {chatMessages.length === 0 && (
                                        <div className="text-muted small text-center py-2">No messages yet.</div>
                                    )}
                                    {chatMessages.map((m, idx) => (
                                        <div key={`${m.timestamp || idx}-${m.user}-${idx}`} className="mb-2">
                                            <div className="d-flex justify-content-between">
                                                <span className="fw-semibold small">{m.user}{m.to && m.to !== 'all' ? ` → @${m.to}` : ''}</span>
                                                <span className="text-muted small">{m.timestamp ? new Date(m.timestamp).toLocaleString() : ''}</span>
                                            </div>
                                            <div className="small">{m.text}</div>
                                        </div>
                                    ))}
                                </div>
                                <div className="card-footer p-2">
                                    <div className="input-group input-group-sm">
                                        <select className="form-select" style={{ maxWidth: 140 }} value={chatRecipient} onChange={(e)=>setChatRecipient(e.target.value)}>
                                            <option value="all">All</option>
                                            {allUsers.map(u => (
                                                <option key={u} value={u}>{u}</option>
                                            ))}
                                        </select>
                                        <input
                                            type="text"
                                            className="form-control"
                                            placeholder="Type a message"
                                            value={chatInput}
                                            onChange={(e) => setChatInput(e.target.value)}
                                            onKeyDown={(e) => {
                                                if (e.key === 'Enter') {
                                                    e.preventDefault();
                                                    if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                                                        const text = chatInput.trim();
                                                        if (text) {
                                                            const to = chatRecipient || 'all';
                                                            ws.current.send(JSON.stringify({ type: 'CHAT_MESSAGE', sheet_id: id, payload: { text, user: username, to } }));
                                                            setChatInput('');
                                                        }
                                                    }
                                                }
                                            }}
                                        />
                                        <button
                                            className="btn btn-outline-primary"
                                            type="button"
                                            onClick={() => {
                                                if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                                                    const text = chatInput.trim();
                                                    if (text) {
                                                        const to = chatRecipient || 'all';
                                                        ws.current.send(JSON.stringify({ type: 'CHAT_MESSAGE', sheet_id: id, payload: { text, user: username, to } }));
                                                        setChatInput('');
                                                    }
                                                }
                                            }}
                                        >Send</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                
            </div>
        </div>
    );
}
