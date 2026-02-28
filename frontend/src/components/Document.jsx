    
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
    Filter,
    Undo2,
    Redo2
} from 'lucide-react';
import { Lock, Code, ChevronDown } from 'lucide-react';
import { isSessionValid, clearAuth, getUsername, authenticatedFetch, apiUrl } from '../utils/auth';
import MarkdownEditorPanel from './MarkdownEditorPanel';
import 'bootstrap/dist/css/bootstrap.min.css';
export default function Document() {
    const navigate = useNavigate();
    const location = useLocation();
    const params = useParams();
    const id = params.id || params.sheetId || '';
    const username = getUsername();

    // Core sheet state
    const [data, setData] = useState({});
    const [auditLog, setAuditLog] = useState([]);
    const [sheetName, setSheetName] = useState('');
    const [projectName, setProjectName] = useState(new URLSearchParams(location.search).get('project') || '');
    const [owner, setOwner] = useState('');
    const [editors, setEditors] = useState([]);
    // Tree structure: row -> parent row number (0 means root)
    const [rowParents, setRowParents] = useState({});

    // UI and editing state
    const [connected, setConnected] = useState(false);
    const [isEditing, setIsEditing] = useState(false);
    const [isDoubleClicked, setIsDoubleClicked] = useState(false);
    const [isSidebarOpen, setSidebarOpen] = useState(false);
    const [isChatOpen, setChatOpen] = useState(true);
    const [showFilters, setShowFilters] = useState(false);
    const [filters, setFilters] = useState({});
    const [sortConfig, setSortConfig] = useState({ col: null, direction: null });
    const [focusedCell, setFocusedCell] = useState({ row: 1, col: 'A' });

    // Grid dimensions and labels
    const ROW_HEADERS = useMemo(() => Array.from({ length: 2000 }, (_, i) => i + 1), []);
    const ROWS = ROW_HEADERS.length;
    const COL_HEADERS = useMemo(() => {
        const letters = [];
        // Generate A-Z, AA-AZ, BA-BZ, ... ZZ (702 columns total)
        for (let i = 0; i < 102; i++) {
            if (i < 26) {
                letters.push(String.fromCharCode(65 + i));
            } else {
                const first = Math.floor((i - 26) / 26);
                const second = (i - 26) % 26;
                letters.push(String.fromCharCode(65 + first) + String.fromCharCode(65 + second));
            }
        }
        return letters;
    }, []);
    const COLS = COL_HEADERS.length;

    const isOwner = username && owner && username === owner;
    const canEdit = !!username && (isOwner || (Array.isArray(editors) && editors.includes(username)));

    const handleUnauthorized = () => {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
    };
    // Row cut/paste state
    const [cutRow, setCutRow] = useState(null);
    // Column cut/paste state
    const [cutCol, setCutCol] = useState(null);

    // Cell style controls
    const [styleBg, setStyleBg] = useState('');
    const [styleBold, setStyleBold] = useState(false);
    const [styleItalic, setStyleItalic] = useState(false);
    // Cell script editor (common)
    const [scriptText, setScriptText] = useState('');

    // Multi-cell selection and clipboard state
    const [selectionStart, setSelectionStart] = useState(null); // { row, col }
    const [selectedRange, setSelectedRange] = useState(null); // Array of { row, col }
    const [isSelecting, setIsSelecting] = useState(false);
    const [isSelectingWithShift, setIsSelectingWithShift] = useState(false);
    const [copiedBlock, setCopiedBlock] = useState(null); // { Rows, Cols, values: string[][], scripts?: string[][] }
    const [contextMenu, setContextMenu] = useState({ visible: false, x: 0, y: 0, cell: null });
    // Chat state
    const [chatMessages, setChatMessages] = useState([]); // [{timestamp, user, text, to}]
    const [chatInput, setChatInput] = useState('');
    const [chatRecipient, setChatRecipient] = useState('all');
    const [allUsers, setAllUsers] = useState([]);
    // Ref for chat panel body
    const chatBodyRef = useRef(null);
    // Ref for script textarea
    const scriptTextareaRef = useRef(null);
    // Refs for option inputs (to track cursor position for range insertion)
    const optionIdRefs = useRef([]);
    const optionDisplayValueRefs = useRef([]);
    const [focusedOptionField, setFocusedOptionField] = useState({ index: null, field: null }); // 'id' or 'displayValue'

    // Scroll chat to bottom on open/init and when chatMessages change
    useEffect(() => {
        if (chatBodyRef.current) {
            chatBodyRef.current.scrollTop = chatBodyRef.current.scrollHeight;
        }
    }, [isSidebarOpen, chatMessages]);
    // Highlight selected audit log entry
    const [selectedAuditId, setSelectedAuditId] = useState(null);
    // Floating diff panel for audit value changes
    const [diffPanel, setDiffPanel] = useState({ visible: false, entry: null, parts: [] });
    // Show/hide system audit logs
    const [showSystemLogs, setShowSystemLogs] = useState(true);
    // Undo/Redo stacks for committed cell value edits
    const [undoStack, setUndoStack] = useState([]); // [{type:'cell_edit', row, col, oldValue, newValue}]
    const [redoStack, setRedoStack] = useState([]);
    // Preserve audit log scroll position across open/close
    const auditLogRef = useRef(null);
    const auditLogScrollTopRef = useRef(0);
    const editingOriginalValueRef = useRef(null);
    const editingOriginalScriptRef = useRef(null);

    // Touch state for swipe gestures
    const [touchStart, setTouchStart] = useState({ x: 0, y: 0 });
    const [touchEnd, setTouchEnd] = useState({ x: 0, y: 0 });

    // Markdown editor panel state (for third column / column C)
    const [mdPanelOpen, setMdPanelOpen] = useState(false);
    const [mdPanelCell, setMdPanelCell] = useState({ row: null, col: null });
    const [mdPanelReadOnly, setMdPanelReadOnly] = useState(false);

    const colIndexMap = useMemo(() => {
        const map = {};
        COL_HEADERS.forEach((c, i) => { map[c] = i; });
        return map;
    }, []);

    const colLabelAt = (index) => COL_HEADERS[index] || null;
    const isCellSelected = (rowLabel, colLabel) => {
        if (!selectedRange || selectedRange.length === 0) return false;
        return selectedRange.some((cell) => cell.row === rowLabel && cell.col === colLabel);
    };

    const startSelection = (rowLabel, colLabel) => {
        if(!connected) return
        //console.log(rowLabel, colLabel);
        setIsSelecting(true);
        setSelectionStart({ row: rowLabel, col: colLabel });
        setSelectedRange([{ row: rowLabel, col: colLabel }]);
        setIsEditing(false);
        setIsDoubleClicked(false);
        setCutRow(null);
        setCutCol(null);
        setContextMenu(prev => ({ ...prev, visible: false }));
    };

    const sendSelection = () => {
         if (!selectedRange || selectedRange.length === 0) return;

        // Derive unique rows in the order they appear in filteredRowHeaders
        const rowIndexMap = new Map(filteredRowHeaders.map((r, i) => [r, i]));
        const uniqueRows = Array.from(new Set(selectedRange.map(c => c.row)))
            .filter(r => rowIndexMap.has(r))
            .sort((a, b) => (rowIndexMap.get(a) ?? 0) - (rowIndexMap.get(b) ?? 0));
        // Columns keep sheet order using colIndexMap
        const uniqueCols = Array.from(new Set(selectedRange.map(c => c.col)))
            .sort((a, b) => (colIndexMap[a] ?? -1) - (colIndexMap[b] ?? -1));

        const values = [];
        const scripts = [];
        const cellTypes = [];
        const optionsArray = [];
        const optionsRangeArray = [];
        const optionSelectedArray = [];
        for (const r of uniqueRows) {
            const rowArr = [];
            const scriptRowArr = [];
            const cellTypeRowArr = [];
            const optionsRowArr = [];
            const optionsRangeRowArr = [];
            const optionSelectedRowArr = [];
            for (const c of uniqueCols) {
                const key = `${r}-${c}`;
                const cell = data[key] || {};
                rowArr.push((cell.value ?? '').toString());
                scriptRowArr.push((cell.script ?? '').toString());
                cellTypeRowArr.push(cell.cell_type ?? 0);
                optionsRowArr.push(cell.options ?? []);
                optionsRangeRowArr.push(cell.options_range ?? '');
                optionSelectedRowArr.push(cell.option_selected ?? []);
            }
            values.push(rowArr);
            scripts.push(scriptRowArr);
            cellTypes.push(cellTypeRowArr);
            optionsArray.push(optionsRowArr);
            optionsRangeArray.push(optionsRangeRowArr);
            optionSelectedArray.push(optionSelectedRowArr);
        }

        const noOfRows = uniqueRows.length;
        const noOfCols = uniqueCols.length;
        const rangeText = getSelectedRange() 
        // Combine projectName, sheet_name, and rangeText in the format {{projectname/sheet_name/rangeText}}
        const combinedRangeText = rangeText ? `${projectName}/${id}/${rangeText}` : '';
        //console.log('Copied selection:',  values);
        // Send selection values to backend so other instances of the same user can paste
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = {
                Rows: noOfRows,
                Cols: noOfCols,
                sheet_name: id,
                values,
                scripts,
                cellTypes,
                optionsArray,
                optionsRangeArray,
                optionSelectedArray,
                rangeText: combinedRangeText,
            };
            ws.current.send(JSON.stringify({ type: 'SELECTION_COPIED', sheet_name: id, payload }));
        }
        setIsSelectingWithShift(false);
    }

    const scrollBeyond = (rowLabel, colLabel) => {
        setTimeout(() => {
            // If extending hits the last visible cell at bottom/right, shift view
            const rowIdx = filteredRowHeaders.indexOf(rowLabel);
            if (rowIdx !== -1) {
                const nfStart = Math.max(rowStart, freezeRowsCount);
                const currentRowEnd = Math.min(nfStart + nonFrozenVisibleRowsCount, filteredRowHeaders.length);
                if (rowIdx + 1 > currentRowEnd - 1) {
                    setRowStart(prev => Math.min(
                        (filteredRowHeaders.length - (nonFrozenVisibleRowsCount) + 1) > 1 ? (filteredRowHeaders.length - (nonFrozenVisibleRowsCount) + 1) : 1,
                        rowStart + 1
                    ));
                }
            }

            const colIdx = COL_HEADERS.indexOf(colLabel);
            if (colIdx !== -1) {
                const currentColNum = colIdx + 1; // 1-based column number
                if (currentColNum > freezeColsCount && currentColNum >= colEnd) {
                    setColStart(prev => Math.min(COLS - nonFrozenVisibleColsCount + 1, prev + 1));
                }
            }
        }, 1000);
    }
    const extendSelectionWithMouse = (rowLabel, colLabel) => {
        if (!isSelecting ) return;
          extendSelection(rowLabel, colLabel);
    };
     const extendSelectionWithShift = (rowLabel, colLabel) => {
        if (!isSelectingWithShift ) return;
          extendSelection(rowLabel, colLabel);
    };

    const extendSelection = (rowLabel, colLabel) => {
        // Allow Shift+Arrow keyboard extension even if not dragging.
        const anchor = selectionStart || (focusedCell && focusedCell.row && focusedCell.col ? focusedCell : null);
        if (!anchor) return;

        // Ensure selection mode remains active while extending
        //if (!isSelecting) setIsSelecting(true);

        scrollBeyond(rowLabel, colLabel);

        // Determine row span based on visual order in filteredRowHeaders
        const startIdx = filteredRowHeaders.indexOf(anchor.row);
        const endIdx = filteredRowHeaders.indexOf(rowLabel);
        if (startIdx === -1 || endIdx === -1) return;
        const from = Math.min(startIdx, endIdx);
        const to = Math.max(startIdx, endIdx);
        const rowsInOrder = filteredRowHeaders.slice(from, to + 1);

        // Compute columns span as before (sheet order)
        const cStartIdx = colIndexMap[anchor.col] ?? -1;
        const cEndIdx = colIndexMap[colLabel] ?? -1;
        const cMin = Math.min(cStartIdx, cEndIdx);
        const cMax = Math.max(cStartIdx, cEndIdx);

        // Build list of selected cells following the current filtered row order
        const cells = [];
        for (const r of rowsInOrder) {
            for (let ci = cMin; ci <= cMax; ci++) {
                const cLbl = colLabelAt(ci);
                if (!cLbl) continue;
                cells.push({ row: r, col: cLbl });
            }
        }
        setSelectedRange(cells);
    };

    const endSelection = () => {
            if(!connected) return
            if (!isSelecting) return;
            setIsSelecting(false);
            //sendSelection();
    };

    const closeContextMenu = () => setContextMenu({ visible: false, x: 0, y: 0, cell: null });

    const showContextMenu = (e, rowLabel, colLabel) => {
        e.preventDefault();
        // Close other popups first
        closeScriptPopup();
        setShowCellTypeDialog(false);
        setIsEditing(false);
        setIsDoubleClicked(false);
        setIsSelectingWithShift(false);
        setContextMenu({ visible: true, x: e.clientX, y: e.clientY, cell: { row: rowLabel, col: colLabel } });
    };

    

    const handlePasteAt = (anchorRow, anchorColLabel) => {
        if (!copiedBlock || !anchorColLabel) return;
        const anchorIdx = colIndexMap[anchorColLabel] ?? -1;
        if (anchorIdx < 0) return;
        const anchorRowIndex = filteredRowHeaders.indexOf(anchorRow);
        if (anchorRowIndex < 0) return;
        const updates = {};
        const scriptUpdates = {};
        const cellTypeUpdates = {};
        const changes = [];
        let hasConflict = false;
        // Prevent pasting into any locked destination cells using filtered row order
        for (let rOff = 0; rOff < copiedBlock.Rows; rOff++) {
            const rIdx = anchorRowIndex + rOff;
            const r = filteredRowHeaders[rIdx];
            if (r == null) break; // stop if we run out of visible filtered rows
            for (let cOff = 0; cOff < copiedBlock.Cols; cOff++) {
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
        // Detect conflicts in destination cells using filtered row order
        for (let rOff = 0; rOff < copiedBlock.Rows; rOff++) {
            const rIdx = anchorRowIndex + rOff;
            const r = filteredRowHeaders[rIdx];
            if (r == null) break;
            for (let cOff = 0; cOff < copiedBlock.Cols; cOff++) {
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

        for (let rOff = 0; rOff < copiedBlock.Rows; rOff++) {
            const rIdx = anchorRowIndex + rOff;
            const r = filteredRowHeaders[rIdx];
            if (r == null) break;
            for (let cOff = 0; cOff < copiedBlock.Cols; cOff++) {
                const cIdx = anchorIdx + cOff;
                if (cIdx < 0 || cIdx >= COLS) continue;
                const cLabel = colLabelAt(cIdx);
                if (!cLabel) continue;
                const key = `${r}-${cLabel}`;
                const value = copiedBlock.values[rOff][cOff] ?? '';
                const scriptVal = copiedBlock.scripts && copiedBlock.scripts[rOff] ? (copiedBlock.scripts[rOff][cOff] ?? '') : '';
                const cellType = copiedBlock.cellTypes && copiedBlock.cellTypes[rOff] ? (copiedBlock.cellTypes[rOff][cOff] ?? 0) : 0;
                const options = copiedBlock.optionsArray && copiedBlock.optionsArray[rOff] ? (copiedBlock.optionsArray[rOff][cOff] ?? []) : [];
                const optionsRange = copiedBlock.optionsRangeArray && copiedBlock.optionsRangeArray[rOff] ? (copiedBlock.optionsRangeArray[rOff][cOff] ?? '') : '';
                const optionSelected = copiedBlock.optionSelectedArray && copiedBlock.optionSelectedArray[rOff] ? (copiedBlock.optionSelectedArray[rOff][cOff] ?? []) : [];
                const oldValue = (data[key]?.value ?? '').toString();
                const oldCellType = data[key]?.cell_type ?? 0;
                const oldOptions = data[key]?.options ?? [];
                const oldOptionsRange = data[key]?.options_range ?? '';
                const oldOptionSelected = data[key]?.option_selected ?? [];
                updates[key] = { value, user: username };
                if (scriptVal && scriptVal.toString().length > 0) {
                    scriptUpdates[key] = { script: scriptVal.toString(), user: username };
                }
                // Check if cell metadata has changed
                const hasCellTypeChange = cellType !== oldCellType || 
                    JSON.stringify(options) !== JSON.stringify(oldOptions) || 
                    optionsRange !== oldOptionsRange;
                if (hasCellTypeChange) {
                    cellTypeUpdates[key] = { 
                        cellType, 
                        options, 
                        optionsRange,
                        oldCellType, 
                        oldOptions, 
                        oldOptionsRange,
                        user: username 
                    };
                }
                changes.push({ 
                    row: String(r), 
                    col: String(cLabel), 
                    oldValue, 
                    newValue: (value ?? '').toString(),
                    oldCellType,
                    newCellType: cellType,
                    oldOptions,
                    newOptions: options,
                    oldOptionsRange,
                    newOptionsRange: optionsRange,
                    oldOptionSelected,
                    newOptionSelected: optionSelected,
                });
            }
        }

        // Apply local state in one batch (preserve existing fields like script/style)
        setData(prev => {
            const next = { ...prev };
            Object.entries(updates).forEach(([key, cell]) => {
                next[key] = { ...(prev[key] || {}), value: cell.value, user: cell.user };
            });
            Object.entries(scriptUpdates).forEach(([key, cell]) => {
                next[key] = { ...(next[key] || prev[key] || {}), script: cell.script, user: cell.user };
            });
            Object.entries(cellTypeUpdates).forEach(([key, cell]) => {
                next[key] = { 
                    ...(next[key] || prev[key] || {}), 
                    cell_type: cell.cellType, 
                    options: cell.options,
                    options_range: cell.optionsRange,
                    user: cell.user 
                };
            });
            return next;
        });

        // Record paste operation for undo/redo as a single batch
        if (changes.length > 0) {
            setUndoStack(prev => [...prev, { type: 'paste', changes }]);
            setRedoStack([]);
        }

        // Broadcast each cell update to server
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            // Send cell type updates first
            Object.entries(cellTypeUpdates).forEach(([key, cell]) => {
                const [rowStr, colLabel] = key.split('-');
                const payload = { 
                    row: rowStr, 
                    col: colLabel, 
                    cell_type: cell.cellType, 
                    options: cell.options,
                    options_range: cell.optionsRange,
                    user: username 
                };
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_TYPE', sheet_name: id, payload }));
            });
            // Send script updates after cell type; backend will execute and broadcast updated values
            Object.entries(scriptUpdates).forEach(([key, cell]) => {
                const [rowStr, colLabel] = key.split('-');
                const payload = { row: rowStr, col: colLabel, script: cell.script, user: username };
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_SCRIPT', sheet_name: id, payload }));
            });
            // Send value updates for cells without scripts in source
            Object.entries(updates).forEach(([key, cell]) => {
                if (scriptUpdates[key]) return; // skip value update if a script will define the value
                const [rowStr, colLabel] = key.split('-');
                const payload = { row: rowStr, col: colLabel, value: cell.value, user: username };
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_name: id, payload }));
            });
        }

        closeContextMenu();
    };

    const ws = useRef(null);

    // Viewport state for virtualized grid
    const [cellModified, setCellModified] = useState(0);
    const [scriptModified, setScriptModified] = useState(0);
    const [rowStart, setRowStart] = useState(1);
    const [visibleRowsCount, setVisibleRowsCount] = useState(15);
    const [colStart, setColStart] = useState(1);
    const [visibleColsCount, setVisibleColsCount] = useState(7);
    // Configurable freeze counts (first N rows, first M columns)
    const freezeRowsCount = 1
    const freezeColsCount = 1
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

    // Floating script editor popup state
    const [scriptPopup, setScriptPopup] = useState({ visible: false, row: null, col: null });
    const [scriptRowSpan, setScriptRowSpan] = useState(1);
    const [scriptColSpan, setScriptColSpan] = useState(1);
    const openScriptPopup = (row, col) => {
        if (!row || !col) return;
        // Close other popups first
        setShowCellTypeDialog(false);
        setContextMenu({ visible: false, x: 0, y: 0, cell: null });
        const key = `${row}-${col}`;
        const existingScript = (data[key]?.script ?? '').toString();
        setScriptText(existingScript);
        editingOriginalScriptRef.current = existingScript;
        const existingRowSpan = parseInt(data[key]?.script_output_row_span ?? 1, 10) || 1;
        const existingColSpan = parseInt(data[key]?.script_output_col_span ?? 1, 10) || 1;
        setScriptRowSpan(existingRowSpan);
        setScriptColSpan(existingColSpan);
        setScriptPopup({ visible: true, row, col });
    };
    const closeScriptPopup = () => setScriptPopup({ visible: false, row: null, col: null });

    const getSelectedRange = () => {
         if (!selectedRange || selectedRange.length === 0) return "";
        
        let rangeText = '';
        if (selectedRange.length === 1) {
            // Single cell: {{A5}}
            const cell = selectedRange[0];
            rangeText = `${cell.col}${cell.row}`;
        } else {
            // Multiple cells: find top-left and bottom-right
            const rows = selectedRange.map(c => c.row).sort((a, b) => a - b);
            const cols = selectedRange.map(c => c.col);
            const colIndices = cols.map(c => colIndexMap[c]).sort((a, b) => a - b);
            
            const topRow = rows[0];
            const bottomRow = rows[rows.length - 1];
            const leftCol = COL_HEADERS[colIndices[0]];
            const rightCol = COL_HEADERS[colIndices[colIndices.length - 1]];
            
            rangeText = `${leftCol}${topRow}:${rightCol}${bottomRow}`;
        }
        return rangeText
    }
    // Insert selected range into script at cursor position
    const insertSelectedRangeIntoScript = () => {
        let rangeText = '';
        if (copiedBlock && copiedBlock.rangeText) {
            rangeText = copiedBlock.rangeText || '';
            copiedBlock.rangeText = ''; // clear after use
        }else if (selectedRange && selectedRange.length > 0) {
            rangeText = getSelectedRange()
        }else  {
            return
        }
        
        
        
        // Insert at cursor position
        const textarea = scriptTextareaRef.current;
        if (textarea) {
            const start = textarea.selectionStart;
            const end = textarea.selectionEnd;
            const currentText = scriptText;
            const newText = currentText.substring(0, start) + "{{"+rangeText + "}}" + currentText.substring(end);
            setScriptText(newText);
            
            // Set cursor position after inserted text
            setTimeout(() => {
                textarea.focus();
                const newCursorPos = start + rangeText.length;
                textarea.setSelectionRange(newCursorPos, newCursorPos);
            }, 0);
        }
    };

   

    const insertRangeIntoOptionsRange = () => {
        let rangeText = '';
        if (copiedBlock && copiedBlock.rangeText) {
            rangeText = copiedBlock.rangeText || '';
            copiedBlock.rangeText = ''; // clear after use
        } else if (selectedRange && selectedRange.length > 0) {
            rangeText = getSelectedRange();
        } else {
            return;
        }
        
        // Update the optionsRange state with the selected range
        setOptionsRange(rangeText);
    };

    // Toggle to show scripts instead of values (read-only mode)
    const [showScripts, setShowScripts] = useState(false);

    // Cell type dialog state
    const [showCellTypeDialog, setShowCellTypeDialog] = useState(false);
    const [cellTypeDialogCell, setCellTypeDialogCell] = useState(null);
    const [selectedCellType, setSelectedCellType] = useState(0); // 0=ValueCell, 1=ScriptCell, 2=ComboBox, 3=MultipleSelection
    const [cellTypeOptions, setCellTypeOptions] = useState([]);
    const [optionsRange, setOptionsRange] = useState(''); // Range like "A1:A10" for dynamic options

    // Option selection dialog state (for ComboBox and MultipleSelection cells)
    const [showOptionDialog, setShowOptionDialog] = useState(false);
    const [optionDialogCell, setOptionDialogCell] = useState(null); // { row, col, cellType, options }
    const [selectedOptions, setSelectedOptions] = useState([]); // Array of selected option indices

    const handleDownloadXlsx = async () => {
        try {
            const projQS = projectName ? `&project=${encodeURIComponent(projectName)}` : '';
            const res = await authenticatedFetch(apiUrl(`/api/export?sheet_name=${encodeURIComponent(id)}${projQS}`), {
                method: 'GET',
            });

            if (res.status === 401) {
                handleUnauthorized();
                return;
            }
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

    // Handler to open cell type dialog
    const openCellTypeDialog = (row, col) => {
        // Close other popups first
        closeScriptPopup();
        setContextMenu({ visible: false, x: 0, y: 0, cell: null });
        const key = `${row}-${col}`;
        const cell = data[key] || {};
        const row2Cell = data[`2-${col}`] || {};
        setCellTypeDialogCell({ row, col });
        setSelectedCellType(row2Cell.cell_type || 0);
        setCellTypeOptions(row2Cell.options || []);
        setOptionsRange(row2Cell.options_range || '');
        setShowCellTypeDialog(true);
    };

    // Handler to save cell type changes
    const handleCellTypeChange = () => {
        if (!cellTypeDialogCell || !ws.current || ws.current.readyState !== WebSocket.OPEN) return;

        const { row, col } = cellTypeDialogCell;
        const payload = {
            //row: String(row), 
            row: "2", // to use same cell type for all rows below 2
            col: String(col),
            cell_type: selectedCellType,
            options: cellTypeOptions,
            options_range: optionsRange,
            user: username
        };

        ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_TYPE', sheet_name: id, payload }));
        setShowCellTypeDialog(false);
        setCellTypeDialogCell(null);
    };

    // Handler to add option to cell type options
    const addCellTypeOption = () => {
        setCellTypeOptions(prev => [...prev, '']);
    };

    // Handler to remove option from cell type options
    const removeCellTypeOption = (index) => {
        setCellTypeOptions(prev => prev.filter((_, i) => i !== index));
    };

    // Handler to update option
    const updateCellTypeOption = (index, value) => {
        setCellTypeOptions(prev => {
            const newOptions = [...prev];
            newOptions[index] = value;
            return newOptions;
        });
    };

    // Open option selection dialog for ComboBox or MultipleSelection cells
    const openOptionDialog = (row, col) => {
        const key = `${row}-${col}`;
        const cell = data[key] || {};
        //use Combox or multipleselection values of row 2 for remaing below rows
        const row2Cell = data[`2-${col}`] || {};
        if (row==1) return;
        if (row2Cell.cell_type !== 2 && row2Cell.cell_type !== 3) return; // Only for ComboBox(2) and MultipleSelection(3)
        
        const options = row2Cell.options || [];
        const currentSelected = cell.option_selected || [];
        
        setOptionDialogCell({ row, col, cellType: row2Cell.cell_type, options });
        setSelectedOptions([...currentSelected]);
        setShowOptionDialog(true);
    };

    // Toggle option selection in the dialog
    const toggleOptionSelection = (index) => {
        if (!optionDialogCell) return;
        
        if (optionDialogCell.cellType === 2) { // ComboBox - single selection
            setSelectedOptions([index]);
        } else if (optionDialogCell.cellType === 3) { // MultipleSelection - multiple
            setSelectedOptions(prev => {
                if (prev.includes(index)) {
                    return prev.filter(i => i !== index);
                } else {
                    return [...prev, index];
                }
            });
        }
    };

    // Save selected options and update cell value
    const saveOptionSelection = () => {
        if (!optionDialogCell || !canEdit) return;
        
        const { row, col, cellType, options } = optionDialogCell;
        const key = `${row}-${col}`;
        const cell = data[key] || {};
        
        // Calculate new value based on selected options
        let newValue = '';
        if (cellType === 2 && selectedOptions.length > 0) { // ComboBox
            newValue = options[selectedOptions[0]] || '';
        } else if (cellType === 3) { // MultipleSelection
            newValue = selectedOptions.map(idx => options[idx] || '').join('; ');
        }
        
        const oldValue = (cell.value ?? '').toString();
        const oldSelected = cell.option_selected || [];
        
        // Only update if changed
        if (newValue === oldValue && JSON.stringify(oldSelected) === JSON.stringify(selectedOptions)) {
            setShowOptionDialog(false);
            setOptionDialogCell(null);
            return;
        }
        
        // Add to undo stack
        setUndoStack(prev => [...prev, {
            type: 'option_selection',
            row: String(row),
            col: String(col),
            oldValue,
            newValue,
            oldSelected,
            newSelected: selectedOptions
        }]);
        setRedoStack([]);
        
        // Update local state
        setData(prev => ({
            ...prev,
            [key]: {
                ...(prev[key] || {}),
                value: newValue,
                option_selected: selectedOptions,
                user: username
            }
        }));
        
        // Send to server
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = {
                row: String(row),
                col: String(col),
                value: newValue,
                option_selected: selectedOptions,
                user: username
            };
            ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_OPTIONS', sheet_name: id, payload }));
        }
        
        // Close dialog
        setShowOptionDialog(false);
        setOptionDialogCell(null);
    };

    // Close option dialog
    const closeOptionDialog = () => {
        setShowOptionDialog(false);
        setOptionDialogCell(null);
        setSelectedOptions([]);
    };

    // Visible ranges accounting for frozen panes
    const nonFrozenVisibleRowsCount = Math.max(0, Math.min(visibleRowsCount, ROWS) - Math.min(freezeRowsCount, ROWS));
    const nonFrozenVisibleColsCount = Math.max(0, Math.min(visibleColsCount, COLS) - Math.min(freezeColsCount, COLS));
    const rowEnd = Math.min(Math.max(rowStart, freezeRowsCount) + nonFrozenVisibleRowsCount - 1, ROWS);
    const colEnd = Math.min(Math.max(colStart, freezeColsCount) + nonFrozenVisibleColsCount - 1, COLS);

    // Filtered rows state
    const [filteredRowHeaders, setFilteredRowHeaders] = useState(ROW_HEADERS);

    const closeAllPopups = () => {
        setContextMenu({ visible: false, x: 0, y: 0, cell: null });
        setShowCellTypeDialog(false);
        closeScriptPopup();
        closeOptionDialog();
    }

    useEffect(() => {
        // Check session validity
        if (!username || !isSessionValid()) {
            handleUnauthorized();
            return;
        }

        // Validate with server token immediately and fetch user preferences
        (async () => {
            try {
                const res = await authenticatedFetch(apiUrl('/api/validate'));
                if (!res.ok) {
                    handleUnauthorized();
                    return;
                }
                // Load user preferences for visible rows/cols
                const prefsRes = await authenticatedFetch(apiUrl('/api/user/preferences'));
                if (prefsRes.ok) {
                    const prefs = await prefsRes.json();
                    if (typeof prefs.visible_rows === 'number' && prefs.visible_rows > 0) {
                        setVisibleRowsCount(Math.min(ROWS, Math.max(1, prefs.visible_rows)));
                        const nonFrozenPrefRows = Math.max(0, Math.min(ROWS, prefs.visible_rows) - Math.min(freezeRowsCount, ROWS));
                        setRowStart((prev) => Math.min(prev, Math.max(Math.min(freezeRowsCount, ROWS) + 1, ROWS - nonFrozenPrefRows + 1)));
                    }
                    if (typeof prefs.visible_cols === 'number' && prefs.visible_cols > 0) {
                        setVisibleColsCount(Math.min(COLS, Math.max(1, prefs.visible_cols)));
                        const nonFrozenPrefCols = Math.max(0, Math.min(COLS, prefs.visible_cols) - Math.min(freezeColsCount, COLS));
                        setColStart((prev) => Math.min(prev, Math.max(1, COLS - nonFrozenPrefCols + 1)));
                    }
                }
            } catch (e) {
                // ignore preference fetch errors
            }
        })();
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
            const projQS = projectName ? `&project=${encodeURIComponent(projectName)}` : '';
            const httpBase = apiUrl('/').replace(/\/$/, '');
            const wsBase = httpBase.replace(/^http/, 'ws');
            const socket = new WebSocket(`${wsBase}/ws?user=${encodeURIComponent(username)}&id=${id}${projQS}`);

            socket.onopen = () => {
                console.log('Connected to WS');
                setConnected(true);

                // SEND initial PING after 5 secs  (connection disconnects in firefox )
                function sendInitialPing() {
                    if (socket.readyState === WebSocket.OPEN) {
                        socket.send(JSON.stringify({ type: 'PING', sheet_name: id }));
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
                    let inString = false;
                    let escapeNext = false;
                    for (let i = 0; i < raw.length; i++) {
                        const ch = raw[i];
                        
                        // Handle escape sequences
                        if (escapeNext) {
                            escapeNext = false;
                            continue;
                        }
                        if (ch === '\\') {
                            escapeNext = true;
                            continue;
                        }
                        
                        // Track string boundaries
                        if (ch === '"') {
                            inString = !inString;
                            continue;
                        }
                        
                        // Only count braces outside of strings
                        if (!inString) {
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
                    } else if (msg.type === 'CHAT_UPDATED') {
                        // Update existing chat message (e.g., read status changed)
                        const updated = msg.payload;
                        if (updated && updated.timestamp) {
                            setChatMessages(prev => prev.map(m => 
                                m.timestamp === updated.timestamp ? updated : m
                            ));
                        }
                    } else if (msg.type === 'SELECTION_SHARED') {
                        const { Rows, Cols, sheet_name, values, scripts, cellTypes, optionsArray, optionsRangeArray, optionSelectedArray, rangeText } = msg.payload || {};
                        if (Rows && Cols && Array.isArray(values)) {
                            setCopiedBlock({ 
                                Rows, 
                                Cols, 
                                values, 
                                scripts: Array.isArray(scripts) ? scripts : undefined,
                                cellTypes: Array.isArray(cellTypes) ? cellTypes : undefined,
                                optionsArray: Array.isArray(optionsArray) ? optionsArray : undefined,
                                optionsRangeArray: Array.isArray(optionsRangeArray) ? optionsRangeArray : undefined,
                                optionSelectedArray: Array.isArray(optionSelectedArray) ? optionSelectedArray : undefined,
                                rangeText,
                            }); 
                        }
                    } else if (msg.type === 'PONG') {
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
                    console.log("Received invalid WS message:", event.data);
                    console.error("WS Parse error", e);
                }
            };

            socket.onclose = () => {
                setConnected(false); setIsEditing(false); setIsDoubleClicked(false);
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
                const res = await authenticatedFetch(apiUrl('/api/users'));
                if (res.status === 401) {
                    handleUnauthorized();
                    return;
                }
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
        setSheetName(id);
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
        // Tree structure: row_parents maps row (string) -> parent row (int)
        if (sheet.row_parents) {
            setRowParents(sheet.row_parents);
        } else {
            setRowParents({});
        }
    };

    const updateCellState = (row, col, value, user) => {
        setData(prev => ({
            ...prev,
            [`${row}-${col}`]: {
                ...(prev[`${row}-${col}`] || {}),
                value,
                user,
            }
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

    const updateCellScriptState = (row, col, script, user, rowSpan = 1, colSpan = 1) => {
        setData(prev => ({
            ...prev,
            [`${row}-${col}`]: {
                ...(prev[`${row}-${col}`] || {}),
                script,
                user,
                script_output_row_span: rowSpan,
                script_output_col_span: colSpan,
            }
        }));
        setScriptModified(1);
    };

    const handleCellChange = (r, c, value) => {
        // Optimistic update
        //updateCellState(String(r), String(c), value, username);
        //send update to server only if changed
        if (cellModified === 0) { return; }
        // Capture undo action if the value changed
        const oldValue = (editingOriginalValueRef.current ?? '').toString();
        const newValue = (value ?? '').toString();
        if (oldValue !== newValue) {
            setUndoStack(prev => [...prev, { type: 'cell_edit', row: String(r), col: String(c), oldValue, newValue }]);
            setRedoStack([]);
        }
        // Send to WB
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const msg = {
                type: 'UPDATE_CELL',
                sheet_name: id,
                payload: { row: String(r), col: String(c), value, user: username }
            };
            ws.current.send(JSON.stringify(msg));
        }
        setCellModified(0);
        editingOriginalValueRef.current = null;
    };

    const handleScriptChange = (r, c, script, rowSpan = 1, colSpan = 1) => {
        //if (scriptModified === 0) { return; } //non-blocking for scripts
        const key = `${r}-${c}`;
        const oldScript = (editingOriginalScriptRef.current ?? (data[key]?.script ?? '')).toString();
        const newScript = (script ?? '').toString();
        if (oldScript !== newScript) {
            setUndoStack(prev => [...prev, { type: 'cell_script', row: String(r), col: String(c), oldValue: oldScript, newValue: newScript }]);
            setRedoStack([]);
        }
        console.log('Submitting script change:', { r, c, script });
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            console.log('WS sending script update');
            const msg = {
                type: 'UPDATE_CELL_SCRIPT',
                sheet_name: id,
                payload: { row: String(r), col: String(c), script, user: username, row_span: Number(rowSpan) || 1, col_span: Number(colSpan) || 1 }
            };
            ws.current.send(JSON.stringify(msg));
        }
        setScriptModified(0);
        editingOriginalScriptRef.current = null;
    };

    const doUndo = () => {
        if (undoStack.length === 0 || !canEdit) return;
        closeAllPopups();
        const last = undoStack[undoStack.length - 1];
        if (last.type === 'cell_edit') {
            const { row, col, oldValue, newValue } = last;
            // Apply local state
            updateCellState(row, col, oldValue, username);
            // Send to server
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_name: id, payload: { row, col, value: oldValue, user: username } }));
            }
            // Move to redo stack
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'paste') {
            const { changes } = last;
            // Apply local state for all changes (restore old values and metadata)
            setData(prev => {
                const next = { ...prev };
                for (const ch of changes) {
                    const key = `${ch.row}-${ch.col}`;
                    next[key] = { 
                        ...(next[key] || {}), 
                        value: ch.oldValue, 
                        cell_type: ch.oldCellType ?? 0,
                        options: ch.oldOptions ?? [],
                        options_range: ch.oldOptionsRange ?? '',
                        option_selected: ch.oldOptionSelected ?? [],
                        user: username 
                    };
                }
                return next;
            });
            // Send revert updates to server for all changes
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                for (const ch of changes) {
                    // Send cell type update if it changed
                    const hasCellTypeChange = (ch.newCellType ?? 0) !== (ch.oldCellType ?? 0) || 
                        JSON.stringify(ch.newOptions ?? []) !== JSON.stringify(ch.oldOptions ?? []) || 
                        (ch.newOptionsRange ?? '') !== (ch.oldOptionsRange ?? '');
                    if (hasCellTypeChange) {
                        ws.current.send(JSON.stringify({ 
                            type: 'UPDATE_CELL_TYPE', 
                            sheet_name: id, 
                            payload: { 
                                row: String(ch.row), 
                                col: String(ch.col), 
                                cell_type: ch.oldCellType ?? 0,
                                options: ch.oldOptions ?? [],
                                options_range: ch.oldOptionsRange ?? '',
                                user: username 
                            } 
                        }));
                    }
                    // Send value update
                    ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_name: id, payload: { row: String(ch.row), col: String(ch.col), value: ch.oldValue, user: username } }));
                }
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'cell_script') {
            const { row, col, oldValue } = last;
            // Apply local state
            updateCellScriptState(row, col, oldValue, username);
            // Send to server
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_SCRIPT', sheet_name: id, payload: { row, col, script: oldValue, user: username} }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'option_selection') {
            const { row, col, oldValue, oldSelected } = last;
            const key = `${row}-${col}`;
            // Apply local state
            setData(prev => ({
                ...prev,
                [key]: {
                    ...(prev[key] || {}),
                    value: oldValue,
                    option_selected: oldSelected,
                    user: username
                }
            }));
            // Send to server
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({
                    type: 'UPDATE_CELL_OPTIONS',
                    sheet_name: id,
                    payload: { row, col, value: oldValue, option_selected: oldSelected, user: username }
                }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'insert_row') {
            const { insertedRow } = last;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'DELETE_ROW', sheet_name: id, payload: { row: String(insertedRow), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_row') {
            const { fromRow, targetRow, destIndex } = last;
            // Inverse move: move the row currently at destIndex back to original fromRow
            const inverseTarget = (fromRow < destIndex) ? (fromRow - 1) : fromRow;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'MOVE_ROW', sheet_name: id, payload: { fromRow: String(destIndex), targetRow: String(inverseTarget), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_row_as_child') {
            const { fromRow, targetRow, destIndex } = last;
            const inverseTarget = (fromRow < destIndex) ? (fromRow - 1) : fromRow;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'MOVE_ROW', sheet_name: id, payload: { fromRow: String(destIndex), targetRow: String(inverseTarget), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'insert_col') {
            const { newCol } = last;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'DELETE_COL', sheet_name: id, payload: { col: String(newCol), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'delete_row') {
            const { row } = last;
            // Re-insert the deleted row at its original position
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'INSERT_ROW', sheet_name: id, payload: { targetRow: String(Number(row) - 1), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'delete_col') {
            const { targetLeft } = last;
            // Re-insert the deleted column to the right of its previous left neighbor (fallback to current first column)
            const targetCol = targetLeft ?? colLabelAt(0);
            if (targetCol && ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'INSERT_COL', sheet_name: id, payload: { targetCol: String(targetCol), user: username } }));
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_col') {
            const { fromCol, targetCol, destIndex } = last;
            const destLabel = colLabelAt(destIndex);
            const origIdx = colIndexMap[String(fromCol)] ?? -1;
            if (destLabel && origIdx >= 0 && ws.current && ws.current.readyState === WebSocket.OPEN) {
                // Compute inverse targetIdx' per MOVE_COL insertion semantics
                const fromIdx2 = destIndex;
                let targetIdxPrime;
                if (fromIdx2 >= origIdx) {
                    targetIdxPrime = origIdx - 1;
                } else {
                    targetIdxPrime = origIdx;
                }
                const targetLabelPrime = colLabelAt(targetIdxPrime);
                if (targetLabelPrime) {
                    ws.current.send(JSON.stringify({ type: 'MOVE_COL', sheet_name: id, payload: { fromCol: String(destLabel), targetCol: String(targetLabelPrime), user: username } }));
                }
            }
            setRedoStack(prev => [...prev, last]);
            setUndoStack(prev => prev.slice(0, -1));
        }
    };

    const doRedo = () => {
        if (redoStack.length === 0 || !canEdit) return;
        closeAllPopups();
        const last = redoStack[redoStack.length - 1];
        if (last.type === 'cell_edit') {
            const { row, col, oldValue, newValue } = last;
            updateCellState(row, col, newValue, username);
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_name: id, payload: { row, col, value: newValue, user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'paste') {
            const { changes } = last;
            // Apply local state for all changes (reapply new values and metadata)
            setData(prev => {
                const next = { ...prev };
                for (const ch of changes) {
                    const key = `${ch.row}-${ch.col}`;
                    next[key] = { 
                        ...(next[key] || {}), 
                        value: ch.newValue, 
                        cell_type: ch.newCellType ?? 0,
                        options: ch.newOptions ?? [],
                        options_range: ch.newOptionsRange ?? '',
                        option_selected: ch.newOptionSelected ?? [],
                        user: username 
                    };
                }
                return next;
            });
            // Send updates to server for all changes
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                for (const ch of changes) {
                    // Send cell type update if it changed
                    const hasCellTypeChange = (ch.newCellType ?? 0) !== (ch.oldCellType ?? 0) || 
                        JSON.stringify(ch.newOptions ?? []) !== JSON.stringify(ch.oldOptions ?? []) || 
                        (ch.newOptionsRange ?? '') !== (ch.oldOptionsRange ?? '');
                    if (hasCellTypeChange) {
                        ws.current.send(JSON.stringify({ 
                            type: 'UPDATE_CELL_TYPE', 
                            sheet_name: id, 
                            payload: { 
                                row: String(ch.row), 
                                col: String(ch.col), 
                                cell_type: ch.newCellType ?? 0,
                                options: ch.newOptions ?? [],
                                options_range: ch.newOptionsRange ?? '',
                                user: username 
                            } 
                        }));
                    }
                    // Send value update
                    ws.current.send(JSON.stringify({ type: 'UPDATE_CELL', sheet_name: id, payload: { row: String(ch.row), col: String(ch.col), value: ch.newValue, user: username } }));
                }
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'cell_script') {
            const { row, col, oldValue, newValue } = last;
            updateCellScriptState(row, col, newValue, username);
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_SCRIPT', sheet_name: id, payload: { row, col, script: newValue, user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'option_selection') {
            const { row, col, newValue, newSelected } = last;
            const key = `${row}-${col}`;
            // Apply local state
            setData(prev => ({
                ...prev,
                [key]: {
                    ...(prev[key] || {}),
                    value: newValue,
                    option_selected: newSelected,
                    user: username
                }
            }));
            // Send to server
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({
                    type: 'UPDATE_CELL_OPTIONS',
                    sheet_name: id,
                    payload: { row, col, value: newValue, option_selected: newSelected, user: username }
                }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'insert_row') {
            const { insertedRow } = last;
            // Re-insert the row at the same position (target = insertedRow - 1)
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'INSERT_ROW', sheet_name: id, payload: { targetRow: String(Number(insertedRow) - 1), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_row') {
            const { fromRow, targetRow, destIndex } = last;
            // Reapply original move
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'MOVE_ROW', sheet_name: id, payload: { fromRow: String(fromRow), targetRow: String(targetRow), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_row_as_child') {
            const { fromRow, targetRow, destIndex } = last;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'MOVE_ROW_AS_CHILD', sheet_name: id, payload: { fromRow: String(fromRow), targetRow: String(targetRow), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'insert_col') {
            const { targetCol } = last;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'INSERT_COL', sheet_name: id, payload: { targetCol: String(targetCol), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'move_col') {
            const { fromCol, targetCol } = last;
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'MOVE_COL', sheet_name: id, payload: { fromCol: String(fromCol), targetCol: String(targetCol), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'delete_row') {
            const { row } = last;
            // Reapply deletion of the row
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'DELETE_ROW', sheet_name: id, payload: { row: String(row), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        } else if (last.type === 'delete_col') {
            const { col } = last;
            // Reapply deletion of the column
            if (ws.current && ws.current.readyState === WebSocket.OPEN) {
                ws.current.send(JSON.stringify({ type: 'DELETE_COL', sheet_name: id, payload: { col: String(col), user: username } }));
            }
            setUndoStack(prev => [...prev, last]);
            setRedoStack(prev => prev.slice(0, -1));
        }
    };

    // Global keyboard shortcuts for undo/redo (when not editing within a cell)
    useEffect(() => {
        const onKeyDown = (e) => {
            const isCtrl = e.ctrlKey || e.metaKey;
            if (!isCtrl) return;
            if (isEditing) return; // avoid interfering with textarea editing
            if (e.key.toLowerCase() === 'z' && !e.shiftKey) {
                e.preventDefault();
                doUndo();
            } else if ((e.key.toLowerCase() === 'y') || (e.key.toLowerCase() === 'z' && e.shiftKey)) {
                e.preventDefault();
                doRedo();
            }
        };
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [isEditing, undoStack, redoStack, canEdit]);

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
                ws.current.send(JSON.stringify({ type: 'RESIZE_COL', sheet_name: id, payload }));
            } else if (type === 'row') {
                const payload = { row: String(label), height: lastSize, user: username };
                ws.current.send(JSON.stringify({ type: 'RESIZE_ROW', sheet_name: id, payload }));
            }
        }
        dragRef.current = { type: null, label: null, startPos: 0, startSize: 0 };
        window.removeEventListener('mousemove', onGlobalMouseMove);
        window.removeEventListener('mouseup', onGlobalMouseUp);
    };

    // Sync toolbar style controls with currently focused cell
    useEffect(() => {
        if (!selectedRange || selectedRange.length === 0) return;
        // use the first cell in selected list
        const first = selectedRange[0];
        if (!first || !first.row || !first.col) return;
        const key = `${first.row}-${first.col}`;
        const cell = data[key] || {};
        setStyleBg(cell.background || '');
        setStyleBold(!!cell.bold);
        setStyleItalic(!!cell.italic);
        //setScriptText(cell.script || '');
        editingOriginalScriptRef.current = (cell.script ?? '').toString();
    }, [ selectedRange, data]);

    const applyStyleToSelectedRange = () => {
        if (!selectedRange || selectedRange.length === 0 || !canEdit) return;
        closeAllPopups();
        for (const sel of selectedRange) {
            const r = sel.row;
            const colLabel = sel.col;
            if (!filteredRowHeaders.includes(r)) continue;
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
                ws.current.send(JSON.stringify({ type: 'UPDATE_CELL_STYLE', sheet_name: id, payload }));
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
        const frozenRows = Array.from({ length: Math.min(freezeRowsCount, ROWS) }, (_, i) => i + 1);
        let newFilteredRowHeaders = activeFilters.length === 0
            ? ROW_HEADERS
            : [
                ...frozenRows,
                ...ROW_HEADERS.filter((rowLabel) => {
                    if (rowLabel <= freezeRowsCount) return false; // avoid duplicate, we add frozen rows explicitly
                    return activeFilters.every(([colLabel, filterVal]) => {
                        const key = `${rowLabel}-${colLabel}`;
                        const cell = data[key] || { value: '' };
                        return String(cell.value).toLowerCase().includes(String(filterVal).toLowerCase());
                    });
                })
            ];

        // Apply sorting if configured
        if (sortConfig && sortConfig.col && sortConfig.direction) {
            const frozenRows = Array.from({ length: Math.min(freezeRowsCount, ROWS) }, (_, i) => i + 1);
            const startIdx = frozenRows.length;
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

            newFilteredRowHeaders = startIdx > 0 ? [...frozenRows, ...rowsToSort] : rowsToSort;
            //console.log("sorted::", newFilteredRowHeaders);
        }

        setFilteredRowHeaders(newFilteredRowHeaders);
    }, [filters, sortConfig, data, freezeRowsCount]);

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

    // Determine if a row is empty (no values, no scripts, no locks)
    const isRowEmpty = (rowLabel) => {
        for (let ci = 0; ci < COLS; ci++) {
            const cLabel = colLabelAt(ci);
            const key = `${rowLabel}-${cLabel}`;
            const cell = data[key];
            if (!cell) continue;
            const hasValue = typeof cell.value === 'string' ? cell.value.trim() !== '' : Boolean(cell.value);
            const hasScript = typeof cell.script === 'string' ? cell.script.trim() !== '' : Boolean(cell.script);
            if (hasValue || hasScript || cell.locked) return false;
        }
        return true;
    };

    // Determine if a column is empty (no values, no scripts, no locks)
    const isColEmpty = (colLabel) => {
        for (let r = 1; r <= ROWS; r++) {
            const key = `${r}-${colLabel}`;
            const cell = data[key];
            if (!cell) continue;
            const hasValue = typeof cell.value === 'string' ? cell.value.trim() !== '' : Boolean(cell.value);
            const hasScript = typeof cell.script === 'string' ? cell.script.trim() !== '' : Boolean(cell.script);
            if (hasValue || hasScript || cell.locked) return false;
        }
        return true;
    };

    const moveCutRowBelow = (targetRow) => {
        if (cutRow == null) return;
        if (isFilterActive) return; // disabled while filters are active

        // Delegate row move to backend; it will broadcast updated sheet
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { fromRow: String(cutRow), targetRow: String(targetRow), user: username };
            // Compute final destination index consistent with backend logic
            let destIndex = Number(targetRow) + 1;
            if (Number(cutRow) < destIndex) destIndex -= 1;
            // Push undo entry for structural move
            setUndoStack(prev => [...prev, { type: 'move_row', fromRow: Number(cutRow), targetRow: Number(targetRow), destIndex }]);
            setRedoStack([]);
            ws.current.send(JSON.stringify({ type: 'MOVE_ROW', sheet_name: id, payload }));
        }

        setCutRow(null);
    };

    const moveCutRowAsChild = (targetRow) => {
        if (cutRow == null) return;
        if (isFilterActive) return;

        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { fromRow: String(cutRow), targetRow: String(targetRow), user: username };
            let destIndex = Number(targetRow) + 1;
            if (Number(cutRow) < destIndex) destIndex -= 1;
            setUndoStack(prev => [...prev, { type: 'move_row_as_child', fromRow: Number(cutRow), targetRow: Number(targetRow), destIndex }]);
            setRedoStack([]);
            ws.current.send(JSON.stringify({ type: 'MOVE_ROW_AS_CHILD', sheet_name: id, payload }));
        }

        setCutRow(null);
    };

    const moveCutColRight = (targetCol) => {
        if (cutCol == null) return;
        if (isFilterActive) return; // keep parity with row behavior
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { fromCol: String(cutCol), targetCol: String(targetCol), user: username };
            // Compute final destination index (0-based) and push undo
            const fromIdx = colIndexMap[String(cutCol)] ?? -1;
            const targetIdx = colIndexMap[String(targetCol)] ?? -1;
            if (fromIdx >= 0 && targetIdx >= 0) {
                let destIdx = targetIdx + 1;
                if (fromIdx < destIdx) destIdx -= 1;
                setUndoStack(prev => [...prev, { type: 'move_col', fromCol: String(cutCol), targetCol: String(targetCol), destIndex: destIdx }]);
                setRedoStack([]);
            }
            ws.current.send(JSON.stringify({ type: 'MOVE_COL', sheet_name: id, payload }));
        }
        setCutCol(null);
    };

    const insertRowBelow = (targetRow) => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetRow: String(targetRow), user: username };
            // Record undo as deletion of the newly inserted row
            const insertedRow = Number(targetRow) + 1;
            setUndoStack(prev => [...prev, { type: 'insert_row', insertedRow }]);
            setRedoStack([]);
            ws.current.send(JSON.stringify({ type: 'INSERT_ROW', sheet_name: id, payload }));
        }
    };

    // Insert a row above the first row (i.e., before row 1)
    const insertRowAboveFirst = () => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetRow: String(0), user: username };
            // The inserted row will be row 1
            setUndoStack(prev => [...prev, { type: 'insert_row', insertedRow: 1 }]);
            setRedoStack([]);
            ws.current.send(JSON.stringify({ type: 'INSERT_ROW', sheet_name: id, payload }));
        }
    };

    // Insert a child row below target row and all its descendants
    const insertChildRow = (targetRow) => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetRow: String(targetRow), user: username };
            ws.current.send(JSON.stringify({ type: 'INSERT_CHILD_ROW', sheet_name: id, payload }));
        }
    };

    // Get depth of a row in tree hierarchy (0 = root)
    const getRowDepth = (row) => {
        let depth = 0;
        let current = Number(row);
        const visited = new Set();
        while (rowParents[String(current)] && rowParents[String(current)] > 0) {
            if (visited.has(current)) break;
            visited.add(current);
            depth++;
            current = rowParents[String(current)];
        }
        return depth;
    };

    // Get all descendants of a row
    const getDescendants = (row) => {
        const result = [];
        const queue = [Number(row)];
        while (queue.length > 0) {
            const current = queue.shift();
            for (const [r, parent] of Object.entries(rowParents)) {
                if (parent === current) {
                    const rNum = Number(r);
                    result.push(rNum);
                    queue.push(rNum);
                }
            }
        }
        return result.sort((a, b) => a - b);
    };

    // Get direct children of a row
    const getDirectChildren = (row) => {
        const children = [];
        for (const [r, parent] of Object.entries(rowParents)) {
            if (parent === Number(row)) {
                children.push(Number(r));
            }
        }
        return children.sort((a, b) => a - b);
    };

    // Check if a row has any children
    const hasChildren = (row) => {
        for (const parent of Object.values(rowParents)) {
            if (parent === Number(row)) return true;
        }
        return false;
    };

    const insertColumnRight = (targetCol) => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { targetCol: String(targetCol), user: username };
            // Compute newly inserted column label (based on current headers)
            const targetIdx = colIndexMap[String(targetCol)] ?? -1;
            if (targetIdx >= 0) {
                const newIdx = targetIdx + 1;
                const newCol = colLabelAt(newIdx);
                if (newCol) {
                    setUndoStack(prev => [...prev, { type: 'insert_col', newCol: String(newCol), targetCol: String(targetCol) }]);
                    setRedoStack([]);
                }
            }
            ws.current.send(JSON.stringify({ type: 'INSERT_COL', sheet_name: id, payload }));
        }
    };

    // Insert a column to the left of the first column (before 'A')
    const insertColumnLeftOfFirst = () => {
        if (isFilterActive) return;
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            // Use empty target to signal left-most insertion to backend
            const payload = { targetCol: String(''), user: username };
            // New column label will be 'A'
            setUndoStack(prev => [...prev, { type: 'insert_col', newCol: String('A'), targetCol: String('') }]);
            setRedoStack([]);
            ws.current.send(JSON.stringify({ type: 'INSERT_COL', sheet_name: id, payload }));
        }
    };

    // Delete row/column only if empty (descendants will be deleted too)
    const deleteRow = (rowLabel) => {
        if (isFilterActive) return;
        const desc = getDescendants(rowLabel);
        const allRows = [Number(rowLabel), ...desc];
        const nonEmptyRows = allRows.filter(r => !isRowEmpty(r));
        if (nonEmptyRows.length > 0) {
            const descendantMsg = desc.length > 0 ? ` and ${desc.length} descendant row(s) (${desc.join(', ')})` : '';
            if (!window.confirm(`Row ${rowLabel}${descendantMsg} contains data. All content will be permanently deleted. Continue?`)) return;
        } else if (desc.length > 0) {
            if (!window.confirm(`This will also delete ${desc.length} descendant row(s): ${desc.join(', ')}. Continue?`)) return;
        }
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { row: String(rowLabel), user: username };
            ws.current.send(JSON.stringify({ type: 'DELETE_ROW', sheet_name: id, payload }));
            // Push undo entry to allow reinsertion at the same index
            setUndoStack(prev => [...prev, { type: 'delete_row', row: Number(rowLabel) }]);
            setRedoStack([]);
        }
    };

    const deleteColumn = (colLabel) => {
        if (isFilterActive) return;
        if (!isColEmpty(colLabel)) {
            if (!window.confirm(`Column ${colLabel} contains data. All content will be permanently deleted. Continue?`)) return;
        }
        if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
            const payload = { col: String(colLabel), user: username };
            ws.current.send(JSON.stringify({ type: 'DELETE_COL', sheet_name: id, payload }));
            // Push undo entry with left-neighbor hint for reinsertion
            const idx = colIndexMap[String(colLabel)] ?? -1;
            const leftLabel = idx > 0 ? colLabelAt(idx - 1) : null;
            setUndoStack(prev => [...prev, { type: 'delete_col', col: String(colLabel), targetLeft: leftLabel }]);
            setRedoStack([]);
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

    const toggleChat = () => {
        setChatOpen(prev => !prev);
    };

    // Navigate to a specific cell and ensure it's visible, then focus it
    const navigateToCell = (targetRow, targetColLabel) => {
        if (!targetRow || !targetColLabel) return;
        // Adjust rowStart only if targetRow is not already within the visible window
        const rowIdx = filteredRowHeaders.indexOf(targetRow);
        if (rowIdx !== -1) {
            const currentStartIdx = Math.max(0, Math.max(rowStart, freezeRowsCount) - 1);
            const currentEndIdx = Math.min(filteredRowHeaders.length - 1, currentStartIdx + nonFrozenVisibleRowsCount - 1);
            const rowVisible = rowIdx >= currentStartIdx && rowIdx <= currentEndIdx;
            if (!rowVisible) {
                if (targetRow <= freezeRowsCount) return; // frozen rows are always visible
                const maxRowStart = Math.max(1, filteredRowHeaders.length - nonFrozenVisibleRowsCount + 1);
                const desiredStart = Math.max(freezeRowsCount + 1, Math.min(maxRowStart, rowIdx + 1));
                setRowStart(desiredStart-1);
            }
        }
        // Adjust colStart only if targetCol is not already within the visible window
        const colIdx = COL_HEADERS.indexOf(targetColLabel);
        if (colIdx !== -1) {
            const colNumber = colIdx + 1; // 1-based
            const colVisible = (colNumber <= freezeColsCount) || (colNumber >= colStart && colNumber <= colEnd);
            if (!colVisible) {
                const maxColStart = Math.max(1, COLS - nonFrozenVisibleColsCount + 1);
                const desiredColStart = Math.max(1, Math.min(maxColStart, colNumber));
                setColStart(desiredColStart-1);
            }
        }
        // Set focus state and focus the element after re-render
        setFocusedCell({ row: targetRow, col: targetColLabel });
        setIsEditing(false);
        setIsDoubleClicked(false);
        setTimeout(() => {
            const el = document.querySelector(`textarea[data-row="${targetRow}"][data-col="${targetColLabel}"]`);
            if (el) {
                el.focus();
                if (typeof el.scrollIntoView === 'function') el.scrollIntoView({ block: 'center', inline: 'center' });
                if (typeof el.select === 'function') el.select();
            }
        }, 50);
    };

    // Navigate from an audit log entry using structured row/col if present, else fallback to details parsing
    const navigateToCellFromDetails = (entryOrDetails) => {
        const entry = (entryOrDetails && typeof entryOrDetails === 'object') ? entryOrDetails : { details: entryOrDetails };
        // Prefer structured coordinates if provided by backend
        if (entry && Number.isInteger(entry.row) && entry.row > 0 && typeof entry.col === 'string' && entry.col) {
            const colLabel = entry.col.toUpperCase();
            if (COL_HEADERS.includes(colLabel)) {
                navigateToCell(entry.row, colLabel);
                return;
            }
        }
        const details = entry.details || '';
        if (!details || typeof details !== 'string') return;
        // Patterns: "Set cell 28,C to ..." or "Changed cell 4,B from ..."
        const match = details.match(/(?:Set|Changed)\s+cell\s+(\d+),([A-Z]+)\s+/i);
        if (match) {
            const row = parseInt(match[1], 10);
            const colLabel = match[2].toUpperCase();
            if (!Number.isNaN(row) && COL_HEADERS.includes(colLabel)) {
                navigateToCell(row, colLabel);
            }
            return;
        }
        // Optional: focus column for resize events like "Set width of column C to 93"
        const colMatch = details.match(/column\s+([A-Z]+)\s+/i);
        if (colMatch) {
            const colLabel = colMatch[1].toUpperCase();
            if (COL_HEADERS.includes(colLabel)) {
                navigateToCell(1, colLabel);
            }
            return;
        }
        // Optional: focus row for resize like "Set height of row 12 to ..."
        const rowMatch = details.match(/row\s+(\d+)\s+/i);
        if (rowMatch) {
            const row = parseInt(rowMatch[1], 10);
            if (!Number.isNaN(row)) {
                navigateToCell(row, focusedCell.col);
            }
        }
    };

    // Compute character-level diff parts between old/new text
    const computeDiffParts = (oldText = '', newText = '') => {
        const a = Array.from(oldText);
        const b = Array.from(newText);
        const n = a.length;
        const m = b.length;
        const dp = Array.from({ length: n + 1 }, () => Array(m + 1).fill(0));
        for (let i = n - 1; i >= 0; i--) {
            for (let j = m - 1; j >= 0; j--) {
                dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
            }
        }
        const parts = [];
        let i = 0, j = 0;
        while (i < n && j < m) {
            if (a[i] === b[j]) {
                let s = '';
                while (i < n && j < m && a[i] === b[j]) { s += a[i]; i++; j++; }
                if (s) parts.push({ type: 'equal', text: s });
            } else if (dp[i + 1][j] >= dp[i][j + 1]) {
                parts.push({ type: 'del', text: a[i] });
                i++;
            } else {
                parts.push({ type: 'add', text: b[j] });
                j++;
            }
        }
        if (i < n) parts.push({ type: 'del', text: a.slice(i).join('') });
        if (j < m) parts.push({ type: 'add', text: b.slice(j).join('') });
        return parts;
    };

    // Open/close diff panel for an audit entry if it has value changes
    const openDiffForEntry = (entry) => {
        if (!entry) { setDiffPanel({ visible: false, entry: null, parts: [] }); return; }
        const oldVal = (entry.old_value ?? '').toString();
        const newVal = (entry.new_value ?? '').toString();
        if (!oldVal  ) { setDiffPanel({ visible: false, entry: null, parts: [] }); return; }
        const parts = computeDiffParts(oldVal, newVal);
        setDiffPanel({ visible: true, entry, parts });
    };
    const closeDiffPanel = () => setDiffPanel({ visible: false, entry: null, parts: [] });

    const nonFrozenRowSliceStart = Math.max(rowStart, freezeRowsCount);
    const nonFrozenRowSliceEnd = Math.min(nonFrozenRowSliceStart + nonFrozenVisibleRowsCount, filteredRowHeaders.length);
    const displayedRowHeaders = [
        ...filteredRowHeaders.slice(0, freezeRowsCount),
        ...filteredRowHeaders.slice(nonFrozenRowSliceStart, nonFrozenRowSliceEnd)
    ];

    const displayedColHeaders = [
        ...COL_HEADERS.slice(0, freezeColsCount),
        ...COL_HEADERS.slice(Math.max(colStart, freezeColsCount), colEnd)
    ];

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
                            className="btn btn-outline-primary btn-sm d-flex align-items-center hover:bg-indigo-100 hover:shadow-md active:bg-indigo-200 active:scale-95 transition-all duration-100"
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
                            <Download className="me-1" />Export
                        </button>
                        <button
                            onClick={() => navigate(projectName ? `/settings/${id}?project=${encodeURIComponent(projectName)}` : `/settings/${id}`)}
                            className="btn btn-outline-primary btn-sm d-flex align-items-center ms-2"
                            title="Settings"
                        >
                            <Settings className="me-1" />Settings
                        </button>
                        <button
                            onClick={toggleSidebar}
                            className={`btn btn-outline-primary btn-sm d-flex align-items-center ${isSidebarOpen ? 'active' : ''}`}
                        >
                            <History className="me-1" />Activity
                        </button>
                        <button
                            onClick={toggleChat}
                            className={`btn btn-outline-primary btn-sm d-flex align-items-center ${isChatOpen ? 'active' : ''}`}
                        >
                            <MessageSquare className="me-1" />Chat
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
                            className="px-3 py-1.5 text-sm rounded border border-gray-300 bg-white hover:bg-indigo-100 hover:shadow-md active:bg-indigo-200 active:scale-95 transition-all duration-100 flex items-center gap-2"
                            onClick={() => setShowFilters(v => !v)}
                            title="Toggle column filters"
                        >
                            <Filter size={16} />
                            {showFilters ? 'Hide Filters' : 'Show Filters'}
                        </button>
                        <button
                            className="px-2 py-1.5 text-sm rounded border border-gray-300 bg-white hover:bg-indigo-100 hover:shadow-md active:bg-indigo-200 active:scale-95 transition-all duration-100 flex items-center gap-1"
                            onClick={doUndo}
                            disabled={!canEdit || undoStack.length === 0}
                            title="Undo (Ctrl+Z)"
                        >
                            <Undo2 size={16} /> Undo
                        </button>
                        <button
                            className="px-2 py-1.5 text-sm rounded border border-gray-300 bg-white hover:bg-gray-100 flex items-center gap-1"
                            onClick={doRedo}
                            disabled={!canEdit || redoStack.length === 0}
                            title="Redo (Ctrl+Y / Ctrl+Shift+Z)"
                        >
                            <Redo2 size={16} /> Redo
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
                                    const nfCount = Math.max(0, val - Math.min(freezeRowsCount, ROWS));
                                    setRowStart((prev) => Math.min(prev, Math.max(Math.min(freezeRowsCount, ROWS) + 1, ROWS - nfCount + 1)));
                                    // persist preference
                                    (async () => {
                                        try {
                                            await authenticatedFetch(apiUrl('/api/user/preferences'), {
                                                method: 'PUT',
                                                headers: { 'Content-Type': 'application/json' },
                                                body: JSON.stringify({ visible_rows: val, visible_cols: visibleColsCount })
                                            });
                                        } catch {}
                                    })();
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
                                    const nfCount = Math.max(0, val - Math.min(freezeColsCount, COLS));
                                    setColStart((prev) => Math.min(prev, Math.max(1, COLS - nfCount + 1)));
                                    // persist preference
                                    (async () => {
                                        try {
                                            await authenticatedFetch(apiUrl('/api/user/preferences'), {
                                                method: 'PUT',
                                                headers: { 'Content-Type': 'application/json' },
                                                body: JSON.stringify({ visible_rows: visibleRowsCount, visible_cols: val })
                                            });
                                        } catch {}
                                    })();
                                }}
                                title="Visible columns"
                            />
                            
                            {/* Style controls for focused cell */}
                        <div className="flex items-center gap-2 ml-2">
                            <span className="text-sm text-gray-600">Bg</span>
                            <input
                                type="color"
                                value={styleBg || '#ffffff'}
                                onChange={(e) => { closeAllPopups(); setStyleBg(e.target.value); }}
                                disabled={!canEdit}
                                title="Background color"
                            />
                            <button
                                className={`px-2 py-1 text-sm rounded border ${styleBold ? 'bg-indigo-100 border-indigo-300' : 'border-gray-300 bg-white'} hover:bg-gray-100`}
                                onClick={() => { closeAllPopups(); setStyleBold(v => !v); }}
                                disabled={!canEdit}
                                title="Bold"
                            >
                                B
                            </button>
                            <button
                                className={`px-2 py-1 text-sm rounded border ${styleItalic ? 'bg-indigo-100 border-indigo-300' : 'border-gray-300 bg-white'} hover:bg-gray-100`}
                                onClick={() => { closeAllPopups(); setStyleItalic(v => !v); }}
                                disabled={!canEdit}
                                title="Italic"
                            >
                                I
                            </button>
                            <button
                                className="px-2 py-1 text-sm rounded border border-gray-300 bg-white hover:bg-indigo-100 hover:shadow-md active:bg-indigo-200 active:scale-95 transition-all duration-100"
                                onClick={applyStyleToSelectedRange}
                                disabled={!canEdit}
                                title="Apply to selected cells"
                            >
                                Apply
                            </button>
                        </div>
                        {/* Show Scripts toggle */}
                        <div className="flex items-center gap-2 ml-4">
                            <label className="inline-flex items-center gap-2 text-sm">
                                <input
                                    type="checkbox"
                                    className="form-check-input"
                                    checked={showScripts}
                                    onChange={(e) => setShowScripts(e.target.checked)}
                                    title="Display cell scripts instead of values (cells become read-only)"
                                />
                                Show Scripts (read-only)
                            </label>
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
                                ←
                            </button>
                        </div>
                        <div className="p-2 border-bottom bg-white">
                            <label className="d-flex align-items-center gap-2 text-sm" style={{ fontSize: '0.875rem', cursor: 'pointer' }}>
                                <input
                                    type="checkbox"
                                    className="form-check-input m-0"
                                    checked={showSystemLogs}
                                    onChange={(e) => setShowSystemLogs(e.target.checked)}
                                    title="Show/hide system audit logs"
                                />
                                Show system logs
                            </label>
                        </div>
                        <div ref={auditLogRef} className="overflow-auto p-3" style={{ height: 'calc(70% - 96px)', overflowY: 'scroll' , opacity: 1 }}>
                            {auditLog.slice().reverse().filter(entry => showSystemLogs || entry.user !== 'system').map((entry, i) => {
                                const ts = entry.timestamp ? new Date(entry.timestamp).toLocaleString() : '';
                                const entryId = `${entry.timestamp || i}|${entry.user || ''}|${entry.action || ''}|${entry.details || ''}`;
                                const isSelected = selectedAuditId === entryId;
                                const canRevert = (
                                    isOwner && 
                                    entry && username !== entry.user && entry.action === 'EDIT_CELL' && entry.user !== 'system' &&
                                    Number.isInteger(entry.row) && entry.row > 0 && typeof entry.col === 'string' && entry.col &&
                                    (data[`${entry.row}-${entry.col}`]?.value ?? '')?.toString() === (entry.new_value ?? '')?.toString() &&
                                    (data[`${entry.row}-${entry.col}`]?.locked !== true)
                                );
                                const onRevert = (e) => {
                                    e.stopPropagation();
                                    if (!canRevert) return;
                                    if (!ws.current || ws.current.readyState !== WebSocket.OPEN) return;
                                    const row = String(entry.row);
                                    const col = String(entry.col);
                                    const oldVal = (entry.old_value ?? '').toString();
                                    ws.current.send(JSON.stringify({
                                        type: 'UPDATE_CELL',
                                        sheet_name: id,
                                        payload: { row, col, value: oldVal, user: username, revert: true }
                                    }));
                                    // Notify entry.user with a chat message
                                    if (entry.user && username !== entry.user) {
                                        const msg = `Your change to cell ${col}${row} was reverted by ${username}.`;
                                        ws.current.send(JSON.stringify({
                                            type: 'CHAT_MESSAGE',
                                            sheet_name: id,
                                            payload: {
                                                text: msg,
                                                user: username,
                                                to: entry.user
                                            }
                                        }));
                                    }
                                };
                                return (
                                    <div
                                        key={entryId}
                                        className={`p-2 mb-2 rounded border opacity-100 `}
                                        onClick={() => { setSelectedAuditId(entryId); navigateToCellFromDetails(entry); openDiffForEntry(entry); }}
                                        title={ts}
                                        style={{ opacity: 1, backgroundColor: entry.change_reversed ? (isSelected ? '#e6e3dc' : '#fee2e2') : (isSelected ? '#beebeb' : '#ffffff') }}
                                    >
                                        <div className="d-flex justify-content-between" style={{ opacity: 1}} >
                                            <span className="fw-semibold small">{entry.user}</span>
                                            <span className="text-muted small">{ts}</span>
                                        </div>
                                        <div className="small d-flex align-items-center justify-content-between">
                                            <div>
                                                <span className="badge bg-light text-dark me-2">{entry.action}</span> {entry.change_reversed ? <del>{entry.details}</del> : entry.details}
                                            </div>
                                            {canRevert && !entry.change_reversed && (
                                                <button className="btn btn-sm btn-outline-danger ms-2" onClick={onRevert} title="Revert to previous value">Revert</button>
                                            )}
                                        </div>
                                    </div>
                                );
                            })}
                            {auditLog.length === 0 && (
                                <div className="text-center text-muted py-5">
                                    <History className="mb-2" size={48} opacity={1} />
                                    <p className="mb-0">No activity yet.</p>
                                </div>
                            )}
                        </div>
                    </div>
                )}
                {diffPanel.visible && (
                    <div style={{ position: 'fixed', right: 392, top: 70, width: 320, zIndex: 1100 }}>
                        <div className="card shadow-sm">
                            <div className="card-header py-2 d-flex align-items-center justify-content-between">
                                <span className="fw-semibold small">Change Preview</span>
                                <button className="btn btn-sm btn-light" onClick={closeDiffPanel} aria-label="Close preview">×</button>
                            </div>
                            <div className="card-body p-2" style={{ maxHeight: 240, overflowY: 'auto' }}>
                                {diffPanel.entry && (
                                    <div className="small mb-2 text-muted">Cell {diffPanel.entry.row ?? ''},{diffPanel.entry.col ?? ''}</div>
                                )}
                                <div className="small" style={{ lineHeight: 1.6 }}>
                                    {diffPanel.parts.map((p, idx) => {
                                        if (p.type === 'add') return <span key={idx} style={{ color: '#16a34a' }}>{p.text}</span>;
                                        if (p.type === 'del') return <span key={idx} style={{ color: '#dc2626', textDecoration: 'line-through' }}>{p.text}</span>;
                                        return <span key={idx}>{p.text}</span>;
                                    })}
                                </div>
                            </div>
                        </div>
                    </div>
                )}
                {scriptPopup.visible && (
                    <div style={{ position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 2100 }} className="bg-white border rounded shadow p-3">
                        <div className="flex items-center gap-2">
                            <span className="text-sm text-gray-600">Python Script [Cell : {String(scriptPopup.col)}{String(scriptPopup.row)}]</span>
                        </div>
                        <div className="flex items-center gap-2">
                            
                            <textarea
                                ref={scriptTextareaRef}
                                className="border rounded px-2 py-1 text-sm"
                                rows={3}
                                style={{ minWidth: 240, resize: 'both', overflow: 'auto' }}
                                value={scriptText}
                                onChange={(e) => setScriptText(e.target.value)}
                                disabled={!canEdit}
                                placeholder={`Edit script for ${String(scriptPopup.col)}${String(scriptPopup.row)}`}
                                title="Edit script"
                            />
                            <div className="flex items-center gap-1 ml-2">
                                <label className="text-xs text-gray-600" title="Rows spanned by script output">Output Rows</label>
                                <input
                                    type="number"
                                    min={1}
                                    className="w-14 border rounded px-2 py-1 text-sm"
                                    value={scriptRowSpan}
                                    onChange={(e) => setScriptRowSpan(Math.max(1, parseInt(e.target.value, 10) || 1))}
                                    disabled={!canEdit}
                                    title="Script output row span"
                                />
                            </div>
                            <div className="flex items-center gap-1 ml-2">
                                <label className="text-xs text-gray-600 ml-2" title="Columns spanned by script output">Output Cols</label>
                                <input
                                    type="number"
                                    min={1}
                                    className="w-14 border rounded px-2 py-1 text-sm"
                                    value={scriptColSpan}
                                    onChange={(e) => setScriptColSpan(Math.max(1, parseInt(e.target.value, 10) || 1))}
                                    disabled={!canEdit}
                                    title="Script output column span"
                                />
                            </div>
                        </div>
                        <div className="mt-2 flex items-center gap-2 justify-between">
                            <button
                                className="px-2 py-1 text-sm rounded border border-gray-300 bg-white hover:bg-green-100 hover:shadow-md active:bg-green-200 active:scale-95 transition-all duration-100"
                                onClick={insertSelectedRangeIntoScript}
                                disabled={!canEdit }
                                title="Insert selected range into script at cursor position"
                            >
                                Insert Range
                            </button>
                            <div className="flex items-center gap-2">
                                <button
                                    className="px-2 py-1 text-sm rounded border border-gray-300 bg-white hover:bg-red-100 hover:shadow-md active:bg-red-200 active:scale-95 transition-all duration-100"
                                    onClick={() => {
                                        const { row, col } = scriptPopup;
                                        const key = row && col ? `${row}-${col}` : null;
                                        const isLocked = key ? (data[key]?.locked === true) : false;
                                        if (!canEdit || isLocked || owner !== username) { closeScriptPopup(); return; }
                                        if (!row || !col) { closeScriptPopup(); return; }
                                        updateCellScriptState(row, col, scriptText, username, scriptRowSpan, scriptColSpan);
                                        handleScriptChange(row, col, scriptText, scriptRowSpan, scriptColSpan);
                                        closeScriptPopup();
                                    }}
                                    disabled={!canEdit || owner !== username || (scriptPopup.row && scriptPopup.col ? (data[`${scriptPopup.row}-${scriptPopup.col}`]?.locked === true) : false)}
                                    title="Apply script (owner only)"
                                >
                                    Apply Script
                                </button>
                                <button
                                    className="px-2 py-1 text-sm rounded border border-gray-300 bg-white hover:bg-indigo-100 hover:shadow-md active:bg-indigo-200 active:scale-95 transition-all duration-100"
                                    onClick={() => { closeScriptPopup();  }}
                                    title="Cancel"
                                >
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                {/* Cell Type Dialog */}
                {showCellTypeDialog && cellTypeDialogCell && (
                    <>
                        
                        {/* Dialog */}
                        <div style={{ 
                            position: 'fixed', 
                            top: '50%', 
                            left: '50%', 
                            transform: 'translate(-50%, -50%)', 
                            zIndex: 2200,
                            backgroundColor: 'white',
                            border: '1px solid #ccc',
                            borderRadius: '8px',
                            boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
                            padding: '20px',
                            minWidth: '400px'
                        }}>
                            <h4 className="mb-3">Change Cell Type for {cellTypeDialogCell.col}{cellTypeDialogCell.row}</h4>
                            
                            <div className="mb-3">
                                <label className="form-label">Cell Type:</label>
                                <select 
                                    className="form-select"
                                    value={selectedCellType}
                                    onChange={(e) => setSelectedCellType(parseInt(e.target.value))}
                                >
                                    <option value={0}>Value Cell</option>
                                    <option value={1}>Script Cell</option>
                                    <option value={2}>ComboBox</option>
                                    <option value={3}>Multiple Selection</option>
                                </select>
                            </div>

                            {(selectedCellType === 2 || selectedCellType === 3) && (
                                <div className="mb-3">
                                   
                                    {cellTypeOptions  && (
                                        <>
                                         <label className="form-label">Options:</label>
                                    <div className="mb-3">
                                        <label className="form-label small">Options Range (e.g., A1:A10):</label>
                                        <input
                                            type="text"
                                            className="form-control form-control-sm"
                                            placeholder="e.g., A1:A10"
                                            value={optionsRange}
                                            onChange={(e) => setOptionsRange(e.target.value)}
                                        />
                                        <small className="text-muted">Specify a range to dynamically populate options from sheet data</small>
                                    </div>
                                        <div className="d-flex gap-2">
                                            <button
                                                className="btn btn-sm btn-outline-secondary"
                                                style={{ fontSize: '0.75rem', padding: '0.15rem 0.4rem' }}
                                                onClick={insertRangeIntoOptionsRange}
                                                title="Insert selected range into options range field"
                                            >
                                                Insert Range
                                            </button>
                                        </div>
                                        </>
                                    )}
                                    <div className="mb-2">
                                        {cellTypeOptions && cellTypeOptions.length > 0 && (optionsRange == '') ? (
                                            cellTypeOptions.map((option, idx) => (
                                                <div key={idx} className="mb-2">
                                                    <div className="d-flex gap-2 mb-1">
                                                        <input
                                                            ref={(el) => optionDisplayValueRefs.current[idx] = el}
                                                            type="text"
                                                            className="form-control form-control-sm"
                                                            placeholder="Option Value"
                                                            value={option || ''}
                                                            onChange={(e) => updateCellTypeOption(idx, e.target.value)}
                                                            onFocus={() => setFocusedOptionField({ index: idx, field: 'option' })}
                                                        />
                                                        <button
                                                            className="btn btn-sm btn-danger"
                                                            onClick={() => removeCellTypeOption(idx)}
                                                        >
                                                            ✕
                                                        </button>
                                                    </div>
                                                    
                                                </div>
                                            ))
                                        ) : (
                                            <p className="text-muted small">No options defined yet</p>
                                        )}
                                    </div>
                                    { (optionsRange == '')  && (
                                    <button
                                        className="btn btn-sm btn-secondary"
                                        onClick={addCellTypeOption}
                                    >
                                        + Add Option
                                    </button>
                                    )}
                                </div>
                            )}

                            <div className="d-flex gap-2 justify-content-end">
                                <button
                                    className="btn btn-primary"
                                    onClick={handleCellTypeChange}
                                >
                                    Save
                                </button>
                                <button
                                    className="btn btn-secondary"
                                    onClick={() => {
                                        setShowCellTypeDialog(false);
                                        setCellTypeDialogCell(null);
                                    }}
                                >
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </>
                )}

                {/* Option Selection Dialog for ComboBox and MultipleSelection */}
                {showOptionDialog && optionDialogCell && (
                    <>
                        <div
                            className="position-fixed top-0 start-0 w-100 h-100 bg-dark bg-opacity-50"
                            style={{ zIndex: 1050 }}
                            onClick={closeOptionDialog}
                        />
                        <div
                            className="position-fixed top-50 start-50 translate-middle bg-white rounded shadow-lg p-4"
                            style={{ zIndex: 1051, minWidth: '400px', maxWidth: '600px', maxHeight: '80vh', overflowY: 'auto' }}
                            onClick={(e) => e.stopPropagation()}
                        >
                            <h5 className="mb-3">
                                {optionDialogCell.cellType === 2 ? 'Select Option (ComboBox)' : 'Select Options (Multiple Selection)'}
                            </h5>
                            <p className="text-muted small mb-3">
                                Cell: {optionDialogCell.row}-{optionDialogCell.col}
                            </p>

                            <div className="mb-3">
                                {optionDialogCell.options && optionDialogCell.options.length > 0 ? (
                                    <div className="list-group">
                                        {optionDialogCell.options.map((option, idx) => (
                                            <button
                                                key={idx}
                                                type="button"
                                                className={`list-group-item list-group-item-action d-flex align-items-center ${
                                                    selectedOptions.includes(idx) ? 'active' : ''
                                                }`}
                                                onClick={() => toggleOptionSelection(idx)}
                                            >
                                                <input
                                                    type={optionDialogCell.cellType === 2 ? 'radio' : 'checkbox'}
                                                    className="form-check-input me-2"
                                                    checked={selectedOptions.includes(idx)}
                                                    readOnly
                                                />
                                                <span>{option}</span>
                                            </button>
                                        ))}
                                    </div>
                                ) : (
                                    <p className="text-muted">No options available</p>
                                )}
                            </div>

                            <div className="d-flex gap-2 justify-content-end">
                                <button
                                    className="btn btn-primary"
                                    onClick={saveOptionSelection}
                                >
                                    Save
                                </button>
                                <button
                                    className="btn btn-secondary"
                                    onClick={closeOptionDialog}
                                >
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </>
                )}

                {/* Grid Area */}
                <div className="flex-1 overflow-hidden p-6 bg-gray-100/50" >
                    {/* Scrollbars + Grid layout */}
                    <div className="h-full w-full" style={{ display: 'grid', gridTemplateColumns: '24px auto', gridTemplateRows: '24px auto 24px' }}>
                        
                        {/* Top horizontal column scrollbar */}
                        <div style={{ gridColumn: '2 / span 1', gridRow: '1 / span 1' }}
                             onWheel={(e) => {
                                 e.preventDefault();
                                 setIsEditing(false);
                                 setIsDoubleClicked(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, COLS - nonFrozenVisibleColsCount + 1);
                                 setColStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}
                             className="flex items-stretch">
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, COLS - nonFrozenVisibleColsCount + 1)}
                                value={colStart}
                                onChange={(e) =>{
                                    setIsEditing(false);
                                    setIsDoubleClicked(false);
                                    const { row, col } = focusedCell;
                                    const key = `${row}-${col}`;
                                    if (data[key]) {
                                        handleCellChange(row, col, data[key].value);
                                    }
                                    setColStart(Math.max(1, Math.min(COLS - nonFrozenVisibleColsCount + 1, parseInt(e.target.value, 10) || 1)))}}
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
                                 const maxStart = Math.max(freezeRowsCount + 1, ROWS - nonFrozenVisibleRowsCount + 1);
                                 setRowStart(prev => Math.max(freezeRowsCount + 1, Math.min(maxStart, prev + step)));
                             }}
                           
                             className="flex items-stretch">
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, ROWS - visibleRowsCount + 1)}
                                value={rowStart}
                                onChange={(e) => {
                                    setIsEditing(false);
                                    setIsDoubleClicked(false);
                                    const { row, col } = focusedCell;
                                    const key = `${row}-${col}`;
                                    if (data[key]) {
                                        handleCellChange(row, col, data[key].value);
                                    }
                                    
                                    setRowStart(Math.max(1, Math.min(ROWS - visibleRowsCount + 1, parseInt(e.target.value, 10) || 1)))}}
                                style={{  writingMode: 'vertical-rl', height:  '100%', width: '100%' }}
                                aria-label="Rows scrollbar"
                            />
                        </div>

                        {/* Grid content */}
                        <div
                            style={{ gridColumn: '2 / span 1', gridRow: '2 / span 1', overflow: 'auto' }}
                            onWheel={(e) => {
                                 // Allow native scrolling; only paginate rows when hitting limits
                                 setIsEditing(false);
                                 setIsDoubleClicked(false);
                                 const { row, col } = focusedCell;
                                 const key = `${row}-${col}`;
                                 if (data[key]) {
                                     handleCellChange(row, col, data[key].value);
                                 }
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, ROWS - visibleRowsCount + 1);
                                // Use browser window scroll position instead of container scroll
                                const winScrollY = typeof window !== 'undefined' ? window.scrollY : 0;
                                const docEl = typeof document !== 'undefined' ? document.documentElement : null;
                                const docScrollTop = docEl ? docEl.scrollTop : 0;
                                const scrollTop = winScrollY || docScrollTop || 0;
                                const viewportBottom = scrollTop + (typeof window !== 'undefined' ? window.innerHeight : 0);
                                const docScrollHeight = docEl ? docEl.scrollHeight : 0;
                                const atTop = scrollTop <= 0;
                                const atBottom = Math.ceil(viewportBottom) >= docScrollHeight;

                                 if ((step < 0 && atTop) || (step > 0 && atBottom)) {
                                   setRowStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                                 }
                             }}
                             onTouchStart={(e) => {
                                 e.preventDefault();
                                 const touch = e.touches[0];
                                 setTouchStart({ x: touch.clientX, y: touch.clientY });
                             }}
                             onTouchMove={(e) => {
                                 e.preventDefault();
                                 const touch = e.touches[0];
                                 setTouchEnd({ x: touch.clientX, y: touch.clientY });
                             }}
                             onTouchEnd={(e) => {
                                 e.preventDefault();
                                 const deltaY = touchStart.y - touchEnd.y;
                                 const deltaX = touchStart.x - touchEnd.x;
                                 const minSwipeDistance = 50;
                                 
                                 // Determine if horizontal or vertical swipe is dominant
                                 const isVerticalSwipe = Math.abs(deltaY) > Math.abs(deltaX);
                                 
                                 if (isVerticalSwipe && Math.abs(deltaY) > minSwipeDistance) {
                                     // Handle vertical swipe (rows)
                                     
                                     setIsEditing(false);
                                     setIsDoubleClicked(false);
                                     const { row, col } = focusedCell;
                                     const key = `${row}-${col}`;
                                     if (data[key]) {
                                         handleCellChange(row, col, data[key].value);
                                     }
                                     const step = deltaY > 0 ? 1 : -1; // swipe up = scroll down, swipe down = scroll up
                                     const maxStart = Math.max(1, ROWS - visibleRowsCount + 1);
                                     // Use browser window scroll position instead of container scroll
                                     const winScrollY = typeof window !== 'undefined' ? window.scrollY : 0;
                                     const docEl = typeof document !== 'undefined' ? document.documentElement : null;
                                     const docScrollTop = docEl ? docEl.scrollTop : 0;
                                     const scrollTop = winScrollY || docScrollTop || 0;
                                     const viewportBottom = scrollTop + (typeof window !== 'undefined' ? window.innerHeight : 0);
                                     const docScrollHeight = docEl ? docEl.scrollHeight : 0;
                                     const atTop = scrollTop <= 0;
                                     const atBottom = Math.ceil(viewportBottom) >= docScrollHeight;

                                     if ((step < 0 && atTop) || (step > 0 && atBottom)) {
                                         setRowStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                                     }
                                 } else if (!isVerticalSwipe && Math.abs(deltaX) > minSwipeDistance) {
                                     // Handle horizontal swipe (columns)
                                     
                                     setIsEditing(false);
                                     setIsDoubleClicked(false);
                                     const { row, col } = focusedCell;
                                     const key = `${row}-${col}`;
                                     if (data[key]) {
                                         handleCellChange(row, col, data[key].value);
                                     }
                                     const step = deltaX > 0 ? 1 : -1; // swipe left = scroll right, swipe right = scroll left
                                     const maxStart = Math.max(1, COLS - nonFrozenVisibleColsCount + 1);
                                     setColStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                                 }
                             }}
                            tabIndex={0}
                            id="grid-container"
                            
                        >
                            <div className="inline-block bg-blue-500 rounded-lg shadow-lg border border-gray-200 overflow-hidden">
                        <table className="border-collapse" >
                            <thead>
                                <tr>
                                    <th
                                        className="bg-gray-50 border-b border-r border-gray-200 p-2 relative select-none"
                                        style={{ width: `${rowLabelWidth}px`, height: `${colHeaderHeight}px` }}
                                    >
                                        <div className="inline-flex items-center gap-1">
                                            <span>➕</span>
                                            {connected && canEdit && (
                                                <button
                                                    type="button"
                                                    className="btn btn-xs btn-light"
                                                    disabled={isFilterActive}
                                                    title={isFilterActive ? 'Disabled while filters are active' : 'Insert column to the left of A'}
                                                    onClick={() => insertColumnLeftOfFirst()}
                                                    style={{ padding: '0 4px', fontSize: '10px' }}
                                                >
                                                    <span role="img" aria-label="insert-col-right">➡️</span>
                                                </button>
                                            )}
                                            {connected && canEdit && (
                                                <button
                                                    type="button"
                                                    className="btn btn-xs btn-light"
                                                    disabled={isFilterActive}
                                                    title={isFilterActive ? 'Disabled while filters are active' : 'Insert row above first row'}
                                                    onClick={() => insertRowAboveFirst()}
                                                    style={{ padding: '0 4px', fontSize: '10px' }}
                                                >
                                                    <span role="img" aria-label="insert-row-below" >⬇️</span>
                                                </button>
                                            )}
                                        </div>
                                    </th>
                                    {displayedColHeaders.map(h => (
                                        <th
                                            key={h}
                                            className="bg-gray-50 border-b border-r border-gray-200 p-2 text-xs font-semibold text-gray-500 uppercase tracking-wider text-center select-none relative"
                                            style={{position: 'relative', width: `${colWidths[h] || DEFAULT_COL_WIDTH}px`, height: `${colHeaderHeight}px` ,padding :`0`}}
                                            onMouseOver={()=> endSelection()}
                                            
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
                                                    {connected && canEdit && (
                                                        <button
                                                            type="button"
                                                            className="btn btn-xs btn-light"
                                                            disabled={isFilterActive}
                                                            title={isFilterActive ? 'Disabled while filters are active' : `Delete column ${h}`}
                                                            onClick={() => deleteColumn(h)}
                                                            style={{ padding: '0 4px', fontSize: '10px' }}
                                                        >
                                                            <span role="img" aria-label="delete-col">🗑️</span>
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
                                                    background: 'rgba(250, 250, 250, 0)', // indigo-500 tint
                                                    borderRight: '1px solid rgb(113, 114, 113)',
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
                                            style={{ width: `${rowLabelWidth}px` , position: 'relative', padding: '0' }}
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
                                            onMouseOver={()=> endSelection()}
                                        >
                                            
                                            {/* Row actions: Insert / Cut / Paste / Insert Child */}
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
                                                {connected && canEdit && (
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Insert child row under ${rowLabel}`}
                                                        onClick={() => insertChildRow(rowLabel)}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="insert-child">🔽</span>
                                                    </button>
                                                )}
                                               
                                                {connected && canEdit && (
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Delete row ${rowLabel}${hasChildren(rowLabel) ? ' and all descendants' : ''}`}
                                                        onClick={() => deleteRow(rowLabel)}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="delete-row">🗑️</span>
                                                    </button>
                                                )}
                                                {cutRow === null && connected && canEdit &&(<button
                                                    type="button"
                                                    className="btn btn-xs btn-light"
                                                    disabled={isFilterActive}
                                                    title={isFilterActive ? 'Disabled while filters are active' : `Cut this row${hasChildren(rowLabel) ? ' and all descendants' : ''}`}
                                                    onClick={() => {
                                                        if (hasLockedInRow(rowLabel)) { alert('Cannot cut: row has locked cell(s).'); return; }
                                                        setCutRow(rowLabel); setCutCol(null);
                                                    }}
                                                    style={{ padding: '0 0px', fontSize: '8px' }}
                                                >
                                                    <span role="img" aria-label="cut">✂️</span>
                                                </button>)}
                                                {cutRow != null && cutRow !== rowLabel && connected && canEdit &&(
                                                    <>
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Paste cut row below row ${rowLabel}`}
                                                        onClick={() => { moveCutRowBelow(rowLabel); setCutRow(null); setCutCol(null); }}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="paste-below">📋↓</span>
                                                    </button>
                                                    <button
                                                        type="button"
                                                        className="btn btn-xs btn-light"
                                                        disabled={isFilterActive}
                                                        title={isFilterActive ? 'Disabled while filters are active' : `Paste cut row as child of row ${rowLabel}`}
                                                        onClick={() => { moveCutRowAsChild(rowLabel); setCutRow(null); setCutCol(null); }}
                                                        style={{ padding: '0 0px', fontSize: '8px' }}
                                                    >
                                                        <span role="img" aria-label="paste-as-child">📋→</span>
                                                    </button>
                                                    </>
                                                )}
                                                
                                            </div>
                                            <span style={{ paddingLeft: `${getRowDepth(rowLabel) * 8}px` }}>
                                                {getRowDepth(rowLabel) > 0 && <span style={{ color: '#9ca3af', fontSize: '8px', marginRight: '2px' }}>{'└'}</span>}
                                                {rowLabel}
                                            </span>
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
                                                    background: 'rgba(247, 247, 248, 0)',
                                                    borderBottom: '1px solid rgb(113, 114, 113)',
                                                    zIndex: 20,
                                                    touchAction: 'none'
                                                }}
                                            ></div>
                                        </td>
                                        {displayedColHeaders.map((colLabel) => {
                                            // Only render cell if sheetId matches current id
                                            
                                            const key = `${rowLabel}-${colLabel}`;
                                            const cell = data[key] || { value: '' };
                                            const row2Cell = data[`2-${colLabel}`] || {};
                                            //const selected = isCellSelected(rowLabel, colLabel);
                                            const inShared = (selectedRange && selectedRange.length > 0) ? (function(){
                                                const rows = selectedRange.map(c => c.row);
                                                const rMin = Math.min(...rows);
                                                const rMax = Math.max(...rows);
                                                const colIdxs = selectedRange.map(c => colIndexMap[c.col] ?? -1);
                                                const cMin = Math.min(...colIdxs);
                                                const cMax = Math.max(...colIdxs);
                                                const cIdx = colIndexMap[colLabel] ?? -1;
                                                return rowLabel >= rMin && rowLabel <= rMax && cIdx >= cMin && cIdx <= cMax;
                                            })() : false;
                                            const boundaryStyles = (function(){
                                                if (!selectedRange || selectedRange.length === 0) return {};
                                                // Use filteredRowHeaders order for row boundaries
                                                const rowIdxs = selectedRange
                                                    .map(c => filteredRowHeaders.indexOf(c.row))
                                                    .filter(i => i !== -1);
                                                if (rowIdxs.length === 0) return {};
                                                const rIdxMin = Math.min(...rowIdxs);
                                                const rIdxMax = Math.max(...rowIdxs);
                                                const curRowIdx = filteredRowHeaders.indexOf(rowLabel);
                                                if (curRowIdx === -1) return {};
                                                // Columns continue to use sheet order via colIndexMap
                                                const colIdxs = selectedRange.map(c => colIndexMap[c.col] ?? -1);
                                                const cMin = Math.min(...colIdxs);
                                                const cMax = Math.max(...colIdxs);
                                                const cIdx = colIndexMap[colLabel] ?? -1;
                                                const color = '#6366f1';
                                                const style = {};
                                                if (curRowIdx === rIdxMin && cIdx >= cMin && cIdx <= cMax) style.borderTop = `2px solid ${color}`;
                                                if (curRowIdx === rIdxMax && cIdx >= cMin && cIdx <= cMax) style.borderBottom = `2px solid ${color}`;
                                                if (cIdx === cMin && curRowIdx >= rIdxMin && curRowIdx <= rIdxMax) style.borderLeft = `2px solid ${color}`;
                                                if (cIdx === cMax && curRowIdx >= rIdxMin && curRowIdx <= rIdxMax) style.borderRight = `2px solid ${color}`;
                                                return style;
                                            })();

                                            return (
                                                <td
                                                    key={key}
                                                    className={`border-b border-r border-gray-200 bg-gray-100 p-0 relative min-w-[7rem] group ${ inShared ? 'bg-indigo-50' : 'bg-white'} hover:bg-green-50/20 transition-colors`}
                                                    style={{ position : 'relative', width: `${colWidths[colLabel] || DEFAULT_COL_WIDTH}px`, height: `${rowHeights[rowLabel] || DEFAULT_ROW_HEIGHT}px`, ...boundaryStyles }}
                                                    onContextMenu={(e) => {  !isEditing && !showScripts && showContextMenu(e, rowLabel, colLabel)}}
                                                >
                                                    <textarea
                                                        className={`w-full h-full px-3 py-1 text-sm outline-none border border-gray-200 focus:border-green-200 focus:ring-0 z-10 relative text-gray-800 resize-none`}
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
                                                        value={showScripts ? (cell.script || '') : (cell.value || '')}

                                                        data-row={rowLabel}
                                                        data-col={colLabel}
                                                        readOnly={showScripts || !!cell.locked || !!cell.script || !canEdit}
                                                        onFocus={() => {
                                                            setFocusedCell({ row: rowLabel, col: colLabel });
                                                            setIsEditing(false);
                                                            setIsDoubleClicked(false);
                                                           
                                                        }}
                                                        onMouseOver={e => { e.target.focus(); }}
                                                        onDoubleClick={(e) => {
                                                            if (showScripts) return;
                                                            if (isEditing) return;
                                                            if (cell.locked || !canEdit) return;
                                                            // If cell at second row is ComboBox or MultipleSelection, open option dialog of 2nd row
                                                            const row2Cell = data[`2-${colLabel}`] || {};
                                                            if (row2Cell.cell_type === 2 || row2Cell.cell_type === 3) {
                                                                e.preventDefault();
                                                                openOptionDialog(rowLabel, colLabel);
                                                                return;
                                                            }
                                                            // If the cell has a script, open the script editor instead of value editing
                                                            if ((cell.script ?? '').toString().length > 0) {
                                                                e.preventDefault();
                                                                openScriptPopup(rowLabel, colLabel);
                                                                // Open markdown editor panel when clicking on third column (C)
                                                                if (colLabel === COL_HEADERS[2]) {
                                                                    setMdPanelCell({ row: rowLabel, col: colLabel });
                                                                    setMdPanelReadOnly(true);
                                                                    setMdPanelOpen(true);
                                                                } else {
                                                                    setMdPanelOpen(false);
                                                                }
                                                                return;
                                                            }
                                                            // Default behavior: enter value edit mode
                                                            e.preventDefault();
                                                            e.target.focus();
                                                            if (connected) {
                                                                closeScriptPopup();
                                                                setShowCellTypeDialog(false);
                                                                setIsEditing(true);
                                                                 // Open markdown editor panel when clicking on third column (C)
                                                                if (colLabel === COL_HEADERS[2]) {
                                                                    setMdPanelCell({ row: rowLabel, col: colLabel });
                                                                    setMdPanelReadOnly(false);
                                                                    setMdPanelOpen(true);
                                                                } else {
                                                                    setMdPanelOpen(false);
                                                                }
                                                                setIsDoubleClicked(true);
                                                                setCutRow(null);
                                                                setCutCol(null);
                                                            }
                                                            editingOriginalValueRef.current = (cell.value ?? '').toString();
                                                            if (typeof e.target.setSelectionRange === 'function') {
                                                                const len = e.target.value.length;
                                                                e.target.setSelectionRange(len, len);
                                                            }
                                                        }}
                                                        onMouseDown={(e) => { 
                                                            if (showScripts) { e.preventDefault(); return; }
                                                            if (isEditing) {
                                                                // In edit mode: allow normal text selection, but keep focus
                                                                // Do NOT call preventDefault or selection handlers
                                                                return;
                                                            }
                                                            e.preventDefault();
                                                            e.target.focus();
                                                            editingOriginalValueRef.current = (cell.value ?? '').toString();
                                                            if (e.button === 0 ) {
                                                                startSelection(rowLabel, colLabel);
                                                            }
                                                        }}
                                                        onMouseEnter={() => { if(!isEditing) extendSelectionWithMouse(rowLabel, colLabel);}   } 
                                                        onMouseUp={(e) => {
                                                                    if(!isEditing) {
                                                                        extendSelectionWithMouse(rowLabel, colLabel);
                                                                        endSelection(); 
                                                                    }
                                                        }}
                                                      
                                                        onKeyDown={(e) => {
                                                            const isMultiline = typeof cell.value === 'string' && cell.value.includes('\n');
                                                            const keys = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter'];
                                                            // Enter edit mode when typing any non-arrow key (including Enter)
                                                            if (!keys.includes(e.key)) {
                                                                if (cell.locked) return;
                                                                // If cell has a script, do not enter value edit mode
                                                                if ((cell.script ?? '').toString().length > 0) return;
                                                                if (connected) { closeAllPopups(); setIsEditing(true); 
                                                                     // Open markdown editor panel when clicking on third column (C)
                                                                    if (colLabel === COL_HEADERS[2]) {
                                                                        setMdPanelCell({ row: rowLabel, col: colLabel });
                                                                        setMdPanelReadOnly(false);
                                                                        setMdPanelOpen(true);
                                                                    } else {
                                                                        setMdPanelOpen(false);
                                                                    }

                                                                }
                                                                return;
                                                            }
                                                            // In edit mode, allow default arrow behavior inside textarea and disable cell navigation if multiline without Shift
                                                            if (isEditing && (isMultiline || (!isMultiline && e.shiftKey)|| isDoubleClicked ) && keys.includes(e.key) ) { return; }
                                                            e.preventDefault();
                                                            let nextRow = rowLabel;
                                                            let nextCol = colLabel;
                                                            const rowIdx = filteredRowHeaders.indexOf(rowLabel);
                                                            const colIdx = COL_HEADERS.indexOf(colLabel);
                                                            
                                                            if (e.key === 'ArrowDown' || e.key === 'Enter') {
                                                                if (rowIdx !== -1 && rowIdx + 1 < filteredRowHeaders.length) {
                                                                    nextRow = filteredRowHeaders[rowIdx + 1];
                                                                    const nfStart = Math.max(rowStart, freezeRowsCount);
                                                                    const currentRowEnd = Math.min(nfStart + nonFrozenVisibleRowsCount, filteredRowHeaders.length );
                                                                    if (rowIdx + 1 > currentRowEnd - 1) {
                                                                        setRowStart(prev => Math.min(filteredRowHeaders.length - (nonFrozenVisibleRowsCount) + 1 > 1 ?  filteredRowHeaders.length - (nonFrozenVisibleRowsCount) + 1 : 1, rowStart+1));
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
                                                                        setColStart(prev => Math.min(COLS - nonFrozenVisibleColsCount + 1, prev + 1));
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
                                                            // If Shift is held, extend the selection from the current cell to the next
                                                            if (e.shiftKey) {
                                                                if (!selectionStart) {
                                                                    // Initialize anchor at current focused cell
                                                                    setSelectionStart({ row: rowLabel, col: colLabel });
                                                                    setSelectedRange([{ row: rowLabel, col: colLabel }]);
                                                                }
                                                                setIsSelectingWithShift(true);
                                                                // Extend using computed next target BEFORE moving focus
                                                                extendSelectionWithShift(nextRow, nextCol);
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
                                                            if (showScripts) return;
                                                            // Update local state for textarea value
                                                            if (cell.locked || (cell.script ?? '').toString().length > 0 || !canEdit) return;
                                                            if (connected)
                                                            updateCellState(rowLabel, colLabel, e.target.value);
                                                            setIsSelecting(false);
                                                            setSelectedRange([]);
                                                            
                                                        }}
                                                        onBlur={(e) => {
                                                            if (showScripts) return;
                                                            setIsEditing(false);
                                                            setIsDoubleClicked(false);
                                                            // Commit value to backend only on blur
                                                            if (!cell.locked && (cell.script ?? '').toString().length === 0 && canEdit) handleCellChange(rowLabel, colLabel, e.target.value);
                                                        }}
                                                    />
                                                    
                                                    {cell.script && (
                                                        <span
                                                            title="Script Cell"
                                                            style={{
                                                                position: 'absolute',
                                                                top: 2,
                                                                left: 2,
                                                                zIndex: 60,
                                                                background: 'rgba(243, 248, 255, 0.95)',
                                                                borderRadius: '4px',
                                                                padding: '0px 0px',
                                                                display: 'inline-flex',
                                                                alignItems: 'center',
                                                                lineHeight: 1,
                                                                boxShadow: '0 0 0 0px rgba(0,0,0,0.05)',
                                                                pointerEvents: 'none'
                                                            }}
                                                        >
                                                            <Code size={12} color="#2563eb" />
                                                        </span>
                                                    )}
                                                    {(row2Cell.cell_type === 2 || row2Cell.cell_type === 3) && rowLabel>1 && (
                                                        <span
                                                            title={row2Cell.cell_type === 2 ? "ComboBox Cell" : "Multiple Selection Cell"}
                                                            style={{
                                                                position: 'absolute',
                                                                top: 2,
                                                                left: 2,
                                                                zIndex: 60,
                                                                background: 'rgba(232, 244, 253, 0.95)',
                                                                borderRadius: '4px',
                                                                padding: '0px 0px',
                                                                display: 'inline-flex',
                                                                alignItems: 'center',
                                                                lineHeight: 1,
                                                                boxShadow: '0 0 0 0px rgba(0,0,0,0.05)',
                                                                pointerEvents: 'none'
                                                            }}
                                                        >
                                                            <ChevronDown size={12} color="#0ea5e9" />
                                                        </span>
                                                    )}
                                                    {cell.locked && (
                                                        <span
                                                            title="Locked"
                                                            style={{
                                                                position: 'absolute',
                                                                top: 2,
                                                                right: 2,
                                                                zIndex: 60,
                                                                background: 'rgba(248, 243, 243, 0.95)',
                                                                borderRadius: '4px',
                                                                padding: '0px 0px',
                                                                display: 'inline-flex',
                                                                alignItems: 'center',
                                                                lineHeight: 1,
                                                                boxShadow: '0 0 0 0px rgba(0,0,0,0.05)',
                                                                pointerEvents: 'none'
                                                            }}
                                                        >
                                                            <Lock size={12} color="#4b5563" />
                                                        </span>
                                                    )}

                                                        {/* Context Menu */}
                                                        {contextMenu.visible && (
                                                            <div
                                                                style={{ position: 'fixed', top: contextMenu.y, left: contextMenu.x, zIndex: 2000, display: 'flex', flexDirection: 'column', minWidth: '180px' }}
                                                                className="bg-white border p-2 text-sm "
                                                                onClick={(e) => e.stopPropagation()}
                                                            >
                                                                {contextMenu.cell && (
                                                                    <>
                                                                        {(() => {
                                                                            // Only show Copy if contextMenu.cell is inside selectedRange
                                                                            let copy = false;
                                                                            if (selectedRange && contextMenu.cell) {
                                                                                copy = selectedRange.some(c => c.row === contextMenu.cell.row && c.col === contextMenu.cell.col);
                                                                            }
                                                                            return copy ? (
                                                                                <button
                                                                                    className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                    onClick={() => {
                                                                                        sendSelection();
                                                                                        closeContextMenu();
                                                                                    }}
                                                                                >
                                                                                    Copy
                                                                                </button>
                                                                                
                                                                            ) : null;
                                                                        })()}
                                                                        </>
                                                                )}
                                                                        <button
                                                                            className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                            disabled={!copiedBlock || !contextMenu.cell}
                                                                            onClick={() => handlePasteAt(contextMenu.cell.row, contextMenu.cell.col)}
                                                                        >
                                                                            Paste
                                                                        </button>
                                                                        <button
                                                                            className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                            disabled={(() => {
                                                                                const key = `${contextMenu.cell?.row}-${contextMenu.cell?.col}`;
                                                                                const cell = data[key] || {};
                                                                                const isLockedBySpan = (cell.locked_by || '').startsWith('script-span ');
                                                                                const isOptionType = cell.cell_type === 2 || cell.cell_type === 3;
                                                                                return !canEdit || !contextMenu.cell || isLockedBySpan || isOptionType;
                                                                            })()}
                                                                            onClick={() => {
                                                                                const key = `${contextMenu.cell?.row}-${contextMenu.cell?.col}`;
                                                                                const cell = data[key] || {};
                                                                                const isLockedBySpan = (cell.locked_by || '').startsWith('script-span ');
                                                                                const isOptionType = cell.cell_type === 2 || cell.cell_type === 3;
                                                                                if (!canEdit || !contextMenu.cell || isLockedBySpan || isOptionType) return;
                                                                                openScriptPopup(contextMenu.cell.row, contextMenu.cell.col);
                                                                                closeContextMenu();
                                                                            }}
                                                                        >
                                                                            Edit Script
                                                                        </button>
                                                                        
                                                                        {isOwner && contextMenu.cell && (
                                                                            <button
                                                                                className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                onClick={() => {
                                                                                    openCellTypeDialog(contextMenu.cell.row, contextMenu.cell.col);
                                                                                    closeContextMenu();
                                                                                }}
                                                                            >
                                                                                Change Cell Type
                                                                            </button>
                                                                        )}
                                                                
                                                                {isOwner && contextMenu.cell && (
                                                                    <>
                                                                        {(() => {
                                                                            // Only show Lock Cell if contextMenu.cell is inside selectedRange
                                                                            let showLock = false;
                                                                            if (selectedRange && contextMenu.cell) {
                                                                                showLock = selectedRange.some(c => c.row === contextMenu.cell.row && c.col === contextMenu.cell.col);
                                                                            }
                                                                            return showLock && !data[`${contextMenu.cell.row}-${contextMenu.cell.col}`]?.locked ? (
                                                                                <button
                                                                                    className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                    onClick={() => {
                                                                                        if (ws.current && ws.current.readyState === WebSocket.OPEN && selectedRange && canEdit) {
                                                                                            for (const sel of selectedRange) {
                                                                                                const r = sel.row;
                                                                                                const colLabel = sel.col;
                                                                                                if (!filteredRowHeaders.includes(r)) continue;
                                                                                                const key = `${r}-${colLabel}`;
                                                                                                if (!data[key]?.locked) {
                                                                                                    const payload = { row: String(r), col: String(colLabel), user: username };
                                                                                                    ws.current.send(JSON.stringify({ type: 'LOCK_CELL', sheet_name: id, payload }));
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                        closeContextMenu();
                                                                                    }}
                                                                                >
                                                                                    Lock Cell
                                                                                </button>
                                                                                
                                                                            ) : null;
                                                                        })()}
                                                                        
                                                                        {(() => {
                                                                            // Only show Unlock Cell if contextMenu.cell is inside selectedRange
                                                                            let showUnlock = false;
                                                                            if (selectedRange && contextMenu.cell) {
                                                                                showUnlock = selectedRange.some(c => c.row === contextMenu.cell.row && c.col === contextMenu.cell.col);
                                                                            }
                                                                            return showUnlock && data[`${contextMenu.cell.row}-${contextMenu.cell.col}`]?.locked && !(data[`${contextMenu.cell.row}-${contextMenu.cell.col}`]?.locked_by?.startsWith('script-span ')) ? (
                                                                                <button
                                                                                    className="block w-full text-left px-2 py-1 hover:bg-gray-100 rounded"
                                                                                    onClick={() => {
                                                                                        if (ws.current && ws.current.readyState === WebSocket.OPEN && selectedRange && canEdit) {
                                                                                            for (const sel of selectedRange) {
                                                                                                const r = sel.row;
                                                                                                const colLabel = sel.col;
                                                                                                if (!filteredRowHeaders.includes(r)) continue;
                                                                                                const key = `${r}-${colLabel}`;
                                                                                                if (data[key]?.locked && !(data[key]?.locked_by?.startsWith('script-span '))) {
                                                                                                    const payload = { row: String(r), col: String(colLabel), user: username };
                                                                                                    ws.current.send(JSON.stringify({ type: 'UNLOCK_CELL', sheet_name: id, payload }));
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                        closeContextMenu();
                                                                                    }}
                                                                                >
                                                                                    Unlock Cell
                                                                                </button>
                                                                            ) : null;
                                                                        })()}
                                                                    </>
                                                                )}
                                                            </div>
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
                                 setIsDoubleClicked(false);
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
                                    setIsDoubleClicked(false);
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

                        {/* Markdown Editor Panel for Column C */}
                        {mdPanelOpen && mdPanelCell.row != null && mdPanelCell.col != null && (
                            <MarkdownEditorPanel
                                cellRow={mdPanelCell.row}
                                cellCol={mdPanelCell.col}
                                value={(data[`${mdPanelCell.row}-${mdPanelCell.col}`] || {}).value || ''}
                                readOnly={mdPanelReadOnly || !canEdit || !!(data[`${mdPanelCell.row}-${mdPanelCell.col}`] || {}).locked}
                                onSave={(newValue) => {
                                    if (canEdit && ws.current && ws.current.readyState === WebSocket.OPEN) {
                                        updateCellState(mdPanelCell.row, mdPanelCell.col, newValue, username);
                                        const msg = {
                                            type: 'UPDATE_CELL',
                                            sheet_name: id,
                                            payload: { row: String(mdPanelCell.row), col: String(mdPanelCell.col), value: newValue, user: username }
                                        };
                                        ws.current.send(JSON.stringify(msg));
                                    }
                                    setMdPanelOpen(false);
                                }}
                                onClose={() => setMdPanelOpen(false)}
                            />
                        )}

                        {/* Chat panel (fixed bottom-right) */}

                        {isChatOpen && (
                        <div style={{ position: 'fixed', right: 16, bottom: 16, width: 360, zIndex: 1100 }}>
                            <div className="card shadow-sm">
                                <div className="card-header py-2 d-flex align-items-center justify-content-between">
                                    <span className="fw-semibold small d-flex align-items-center"><MessageSquare size={16} className="me-2"/> Chat</span>
                                    <span className="badge bg-light text-dark">{chatMessages.length}</span>
                                </div>
                                <div className="card-body p-2" style={{ maxHeight: 240, overflowY: 'auto' }}>
                                    <div ref={chatBodyRef} style={{ maxHeight: 240, overflowY: 'auto' }}>
                                        {chatMessages.length === 0 && (
                                            <div className="text-muted small text-center py-2">No messages yet.</div>
                                        )}
                                        {chatMessages.map((m, idx) => {
                                            const isRead = m.read_by && m.read_by[username];
                                            const hasSheetInfo = m.sheet_name;
                                            return (
                                            <div 
                                                key={`${m.timestamp || idx}-${m.user}-${idx}`} 
                                                className={`mb-2 d-flex ${m.user === username ? 'justify-content-end' : 'justify-content-start'}`}
                                            >
                                                <div 
                                                    style={{ 
                                                        maxWidth: '80%', 
                                                        background: m.user === username ? '#f7f5af' : '#f3f4f6', 
                                                        borderRadius: '12px', 
                                                        padding: '8px 12px',
                                                        boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
                                                        cursor: hasSheetInfo ? 'pointer' : 'default'
                                                    }}
                                                    onClick={async () => {
                                                        if (hasSheetInfo && m.sheet_name && m.project_name) {
                                                            // Verify sheet exists before navigation
                                                            try {
                                                                const projQS = m.project_name ? `?project=${encodeURIComponent(m.project_name)}` : '';
                                                                const res = await authenticatedFetch(apiUrl(`/api/sheet/${encodeURIComponent(m.sheet_name)}${projQS}`));
                                                                
                                                                if (res.status === 404) {
                                                                    alert(`Sheet "${m.sheet_name}" no longer exists.`);
                                                                    return;
                                                                }
                                                                
                                                                if (res.status === 401) {
                                                                    handleUnauthorized();
                                                                    return;
                                                                }
                                                                
                                                                if (!res.ok) {
                                                                    alert('Unable to access sheet. Please try again.');
                                                                    return;
                                                                }
                                                                
                                                                // Mark as read
                                                                if (ws.current && ws.current.readyState === WebSocket.OPEN && !isRead) {
                                                                    ws.current.send(JSON.stringify({ 
                                                                        type: 'CHAT_READ', 
                                                                        sheet_name: id, 
                                                                        payload: { timestamp: m.timestamp, user: username } 
                                                                    }));
                                                                }
                                                                
                                                                // Navigate to the sheet
                                                                navigate(`/sheet/${m.sheet_name}?project=${encodeURIComponent(m.project_name)}`);
                                                            } catch (err) {
                                                                console.error('Error verifying sheet:', err);
                                                                alert('Unable to verify sheet existence. Please try again.');
                                                            }
                                                        }
                                                    }}
                                                >
                                                    <div className="d-flex justify-content-between align-items-center">
                                                        <span className="fw-semibold small">
                                                            {m.user === username ? 'me' : m.user}
                                                            {m.to && m.to !== 'all' ? ` → ${m.to === username ? 'me' : m.to}` : ''}
                                                        </span>
                                                        <span className="text-muted" style={{ fontSize: '0.75em', marginLeft: 8 }}>
                                                            {m.timestamp ? new Date(m.timestamp).toLocaleString() : ''}
                                                        </span>
                                                    </div>
                                                    <div className="small">{m.text}</div>
                                                    {hasSheetInfo && (
                                                        <div className="text-muted" style={{ fontSize: '0.7em', marginTop: 4, fontStyle: 'italic' }}>
                                                            📄 {m.sheet_name || 'Untitled'} ({m.project_name || 'Unknown Project'}){isRead ? '    ✓' : ''}
                                                        </div>
                                                    )}
                                                    
                                                    
                                                </div>
                                            </div>
                                        );})}
                                    </div>
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
                                                            ws.current.send(JSON.stringify({ type: 'CHAT_MESSAGE', sheet_name: id, payload: { text, user: username, to } }));
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
                                                        ws.current.send(JSON.stringify({ type: 'CHAT_MESSAGE', sheet_name: id, payload: { text, user: username, to } }));
                                                        setChatInput('');
                                                    }
                                                }
                                            }}
                                        >Send</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                        )}
                    </div>
                </div>

                
            </div>
        </div>
    );
}
