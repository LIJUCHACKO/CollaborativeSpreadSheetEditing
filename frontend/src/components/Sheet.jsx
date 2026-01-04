import React, { useEffect, useState, useRef } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import {
    FileSpreadsheet,
    ArrowLeft,
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
import { isSessionValid, clearAuth, getUsername } from '../utils/auth';
import './bootstrap/dist/css/bootstrap.min.css';

const ROWS = 500;
const COLS = 50;
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
    // Column filters state
    const [filters, setFilters] = useState({});
    const [showFilters, setShowFilters] = useState(false);
    const [focusedCell, setFocusedCell] = useState({ row: 1, col: COL_HEADERS[0] });
    const [isEditing, setIsEditing] = useState(false);

    const ws = useRef(null);

    // Viewport state for virtualized grid
    const [rowStart, setRowStart] = useState(1);
    const [visibleRowsCount, setVisibleRowsCount] = useState(15);
    const [colStart, setColStart] = useState(1);
    const [visibleColsCount, setVisibleColsCount] = useState(7);
    const DEFAULT_COL_WIDTH = 112; // px (Tailwind w-28)
    const DEFAULT_ROW_HEIGHT = 40; // px (Tailwind h-10)
    const DEFAULT_ROW_LABEL_WIDTH = 40; // px (Tailwind w-10)
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

        // Connect to WS
        const socket = new WebSocket(`ws://localhost:8080/ws?user=${encodeURIComponent(username)}&id=${id}`);

        socket.onopen = () => {
            console.log('Connected to WS');
            setConnected(true);
        };

        socket.onmessage = (event) => {
            try {
                const msg = JSON.parse(event.data);
                if (msg.type === 'INIT') {
                    setInitialState(msg.payload);
                } else if (msg.type === 'UPDATE_CELL') {
                    const { row, col, value, user } = msg.payload;
                    updateCellState(row, col, value, user);
                }
            } catch (e) {
                console.error("WS Parse error", e);
            }
        };

        socket.onclose = () => setConnected(false);

        ws.current = socket;

        return () => {
            socket.close();
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
    };

    const updateCellState = (row, col, value, user) => {
        setData(prev => ({
            ...prev,
            [`${row}-${col}`]: { value, user }
        }));

        setAuditLog(prev => [...prev, {
            timestamp: new Date().toISOString(),
            user: user,
            action: "EDIT_CELL",
            details: `Set cell ${row},${col} to ${value}`
        }]);
    };

    const handleCellChange = (r, c, value) => {
        // Optimistic update
        updateCellState(String(r), String(c), value, username);

        // Send to WB
        if (ws.current && ws.current.readyState === WebSocket.OPEN) {
            const msg = {
                type: 'UPDATE_CELL',
                sheet_id: id,
                payload: { row: String(r), col: String(c), value, user: username }
            };
            ws.current.send(JSON.stringify(msg));
        }
    };

    const onGlobalMouseMove = (e) => {
        const { type, label, startPos, startSize } = dragRef.current || {};
        if (!type) return;
        if (type === 'col') {
            const delta = e.clientX - startPos;
            const newSize = Math.max(40, startSize + delta);
            setColWidths(prev => ({ ...prev, [label]: newSize }));
        } else if (type === 'row') {
            const delta = e.clientY - startPos;
            const newSize = Math.max(24, startSize + delta);
            setRowHeights(prev => ({ ...prev, [label]: newSize }));
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
        dragRef.current = { type: null, label: null, startPos: 0, startSize: 0 };
        window.removeEventListener('mousemove', onGlobalMouseMove);
        window.removeEventListener('mouseup', onGlobalMouseUp);
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

    // Update filteredRowHeaders when filters change
    useEffect(() => {
        const activeFilters = Object.entries(filters).filter(([col, val]) => val && val.trim() !== '');
        const newFilteredRowHeaders = activeFilters.length === 0
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
        setFilteredRowHeaders(newFilteredRowHeaders);
    }, [filters]);
    // Determine RowStartfromFilter based on filteredRowHeaders
    //console.log(filteredRowHeaders)
    const RowStartfromFilter = filteredRowHeaders.includes(rowStart +1 )
        ? rowStart +1
        : filteredRowHeaders.find((row) => row > rowStart+1) || filteredRowHeaders[filteredRowHeaders.length -1];
    //console.log("::RowStartfromFilter", RowStartfromFilter);
    const filterstartIndex = filteredRowHeaders.indexOf(RowStartfromFilter);
    const filterstartIndexNew = filterstartIndex + visibleRowsCount  > filteredRowHeaders.length ? filteredRowHeaders.length-visibleRowsCount:filterstartIndex ; 
    //console.log("filterstartIndexNew", filteredRowHeaders[filterstartIndexNew]);
    const displayedRowHeaders = [
        1,
        ...filteredRowHeaders.slice(
            filteredRowHeaders.length > visibleRowsCount?  filterstartIndexNew:1,
            Math.min(filterstartIndexNew + visibleRowsCount , filteredRowHeaders.length )
        )
    ];

    const displayedColHeaders = [COL_HEADERS[0], ...COL_HEADERS.slice(colStart , colEnd)];

    // Clear filter values when showFilters is set to false
    useEffect(() => {
        if (!showFilters) {
            setFilters({});
        }
    }, [showFilters]);

    return (
        <div className="flex h-screen flex-col bg-gray-50 overflow-hidden font-sans text-gray-900">
            {/* Bootstrap Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                    <button
                            onClick={() => navigate('/dashboard')}
                            className="btn btn-outline-primary btn-sm d-flex align-items-center"
                        >
                            <ArrowLeft className="me-1" />
                        </button>
                    <span className="navbar-text d-flex align-items-center fw-bold ">
                        <FileSpreadsheet className="me-2" />{sheetName}
                        <span className="mx-3">|</span>
                        <button className="btn btn-outline-primary btn-sm d-flex align-items-center"><Download className="me-1" /></button>
                        <button
                            onClick={() => setSidebarOpen(!isSidebarOpen)}
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
                        
                    </div>
                </div>
            </header>

            <div style={{ display: 'inline' ,float: 'left'}} className="flex flex-1 overflow-hidden relative">
                {/* Sidebar / Audit Log */}
                {isSidebarOpen && (
                    <div
                        className="position-absolute top-0 end-0 bg-white border-start shadow-lg"
                        style={{ width: '500px',  zIndex: 1050, height: '100%' }}
                    >
                        <div className="d-flex justify-content-between align-items-center p-3 border-bottom bg-light">
                            <h5 className="mb-0 d-flex align-items-center">
                                <History className="me-2" size={18} /> Activity Log
                            </h5>
                            <button 
                                onClick={() => setSidebarOpen(false)} 
                                className="btn btn-sm btn-light"
                                aria-label="Close sidebar"
                            >
                                <ArrowLeft size={18} />
                            </button>
                        </div>
                        <div className="overflow-auto p-3" style={{ height: 'calc(100% - 56px)' ,overflowY:'scroll'}}>
                            {auditLog.slice().reverse().map((entry, i) => (
                                <div key={i} className="d-flex gap-2 mb-3 p-2 rounded hover-bg-light">
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
                            ))}
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
                    <div className="h-full w-full" style={{ display: 'grid', gridTemplateColumns: '24px auto', gridTemplateRows: '24px auto' }}>
                        {/* Top horizontal column scrollbar */}
                        <div style={{ gridColumn: '2 / 3', gridRow: '1 / 2' }}
                             onWheel={(e) => {
                                 e.preventDefault();
                                 const step = e.deltaY > 0 ? 1 : -1;
                                 const maxStart = Math.max(1, COLS - visibleColsCount + 1);
                                 setColStart(prev => Math.max(1, Math.min(maxStart, prev + step)));
                             }}>
                            <input
                                type="range"
                                min={1}
                                max={Math.max(1, COLS - visibleColsCount + 1)}
                                value={colStart}
                                onChange={(e) => setColStart(Math.max(1, Math.min(COLS - visibleColsCount + 1, parseInt(e.target.value, 10) || 1)))}
                                style={{ width: '100%' }}
                                aria-label="Columns scrollbar"
                            />
                        </div>

                        {/* Left vertical row scrollbar */}
                        <div style={{ gridColumn: '1 / 2', gridRow: '2 / 3' }}
                             onWheel={(e) => {
                                 e.preventDefault();
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
                                onChange={(e) => setRowStart(Math.max(1, Math.min(ROWS - visibleRowsCount + 1, parseInt(e.target.value, 10) || 1)))}
                                style={{ writingMode: 'vertical-rl', height: '100%', width: '100%' }}
                                aria-label="Rows scrollbar"
                            />
                        </div>

                        {/* Grid content */}
                        <div style={{ gridColumn: '2 / 3', gridRow: '2 / 3', overflow: 'hidden' }}>
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
                                            style={{ width: `${colWidths[h] || DEFAULT_COL_WIDTH}px`, height: `${colHeaderHeight}px` }}
                                        >
                                            {h}  
                                            <span
                                                onMouseDown={(e) => onColResizeMouseDown(h, e)}
                                                title="Drag to resize column"
                                                role="separator"
                                                aria-orientation="vertical"
                                                style={{
                                                    position: 'relative',
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
                                                className="bg-gray-50 border-b border-r border-gray-200 p-1"
                                                style={{ width: `${colWidths[h] || DEFAULT_COL_WIDTH}px` }}
                                            >
                                                <input
                                                    type="text"
                                                    className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:border-indigo-500"
                                                    placeholder="Filter"
                                                    value={filters[h] || ''}
                                                    onChange={(e) => setFilters(prev => ({ ...prev, [h]: e.target.value }))}
                                                />
                                            </th>
                                        ))}
                                    </tr>
                                )}
                            </thead>
                            <tbody>
                                {displayedRowHeaders.map((rowLabel) => (
                                    <tr key={rowLabel}>
                                        <td
                                            className="bg-gray-50 border-b border-r border-gray-200 p-2 text-center text-xs font-semibold text-gray-500 select-none relative"
                                            style={{ height: `${rowHeights[rowLabel] || DEFAULT_ROW_HEIGHT}px`, width: `${rowLabelWidth}px` }}
                                        >
                                            {rowLabel}
                                             <div
                                                onMouseDown={(e) => onRowResizeMouseDown(rowLabel, e)}
                                                title="Drag to resize row"
                                                role="separator"
                                                aria-orientation="horizontal"
                                                style={{
                                                    position: 'relative',
                                                    left: 0,
                                                    bottom: 0,
                                                    width: '100%',
                                                    height: '8px',
                                                    cursor: 'row-resize',
                                                    userSelect: 'none',
                                                    background: 'rgba(99,102,241,0.15)',
                                                    borderTop: '1px solid #6366f1',
                                                    zIndex: 20,
                                                    touchAction: 'none'
                                                }}
                                            ></div>
                                        </td>
                                        {displayedColHeaders.map((colLabel) => {
                                            const key = `${rowLabel}-${colLabel}`;
                                            const cell = data[key] || { value: '' };
                                            return (
                                                <td
                                                    key={key}
                                                    className="border-b border-r border-gray-200 p-0 relative min-w-[7rem] group bg-white hover:bg-indigo-50/20 transition-colors"
                                                    style={{ width: `${colWidths[colLabel] || DEFAULT_COL_WIDTH}px`, height: `${rowHeights[rowLabel] || DEFAULT_ROW_HEIGHT}px` }}
                                                >
                                                    <input
                                                        className="w-full h-full px-3 py-1 text-sm outline-none border-2 border-transparent focus:border-indigo-500 focus:ring-0 z-10 relative bg-transparent text-gray-800"
                                                        type="text"
                                                        style={{ width: '100%', height: '100%', boxSizing: 'border-box', display: 'block' }}
                                                        value={cell.value}
                                                        data-row={rowLabel}
                                                        data-col={colLabel}
                                                        onFocus={() => { setFocusedCell({ row: rowLabel, col: colLabel }); setIsEditing(false); }}
                                                        onDoubleClick={(e) => { setIsEditing(true); if (typeof e.target.select === 'function') e.target.select(); }}
                                                        onBlur={() => setIsEditing(false)}
                                                        onKeyDown={(e) => {
                                                            const keys = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'];
                                                            // Enter edit mode when typing any non-arrow key
                                                            if (!keys.includes(e.key)) { setIsEditing(true); return; }
                                                            // In edit mode, allow default left/right behavior and disable navigation
                                                            if (isEditing && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) { return; }
                                                            e.preventDefault();
                                                            let nextRow = rowLabel;
                                                            let nextCol = colLabel;
                                                            const rowIdx = filteredRowHeaders.indexOf(rowLabel);
                                                            const colIdx = COL_HEADERS.indexOf(colLabel);
                                                            if (e.key === 'ArrowDown') {
                                                                if (rowIdx !== -1 && rowIdx + 1 < filteredRowHeaders.length) {
                                                                    nextRow = filteredRowHeaders[rowIdx + 1];
                                                                    const currentRowEnd = Math.min(filterstartIndexNew + visibleRowsCount , filteredRowHeaders.length );
                                                                    //console.log("currentRowEnd", currentRowEnd);
                                                                    if (nextRow > filteredRowHeaders[currentRowEnd - 1]) {
        
                                                                        setRowStart(prev => Math.min(ROWS - visibleRowsCount + 1, filteredRowHeaders[filterstartIndexNew ]));
                                                                    }
                                                                }
                                                            } else if (e.key === 'ArrowUp') {
                                                                if (rowIdx > 0) {
                                                                    nextRow = filteredRowHeaders[rowIdx - 1];
                                                                    if (nextRow < rowStart + 1 && nextRow > 1) {
                                                                        setRowStart(prev => Math.max(1, filteredRowHeaders[rowIdx - 2]));
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
                                                                //console.log("colIdx", colIdx);
                                                                if (colIdx > 0) {
                                                                    nextCol = COL_HEADERS[colIdx - 1];
                                                                    const nextColNum = colIdx;
                                                                    if (nextColNum <= colStart) {
                                                                        setColStart(prev => Math.max(1, prev - 1));
                                                                    }
                                                                    //console.log("nextCol", nextCol);
                                                                    //console.log("colStart", colStart);
                                                                    
                                                                }
                                                            }
                                                            setFocusedCell({ row: nextRow, col: nextCol });
                                                            setTimeout(() => {
                                                                const el = document.querySelector(`input[data-row="${nextRow}"][data-col="${nextCol}"]`);
                                                                if (el) {
                                                                    el.focus();
                                                                    if (typeof el.select === 'function') el.select();
                                                                }
                                                            }, 0);
                                                            }}
                                                        onChange={(e) => handleCellChange(rowLabel, colLabel, e.target.value)}
                                                    />
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
                    </div>
                </div>

                
            </div>
        </div>
    );
}
