import React, { useState, useRef, useEffect, useCallback } from 'react';
import { X, Maximize2, Minimize2, Code, Play, CornerDownLeft, Eye } from 'lucide-react';
import { authenticatedFetch, apiUrl } from '../utils/auth';

/**
 * Python keywords that trigger indentation on the next line when a line ends with ':'
 */
const PYTHON_KEYWORDS = [
    'def', 'class', 'if', 'elif', 'else', 'for', 'while', 'with', 'try',
    'except', 'finally', 'async', 'match', 'case'
];

/**
 * Simple Python syntax highlighting (returns array of spans).
 * Handles keywords, strings, comments, numbers, decorators, and builtins.
 */
function highlightPython(code) {
    const lines = code.split('\n');
    return lines.map((line, li) => {
        const tokens = [];
        let i = 0;
        while (i < line.length) {
            // Comments
            if (line[i] === '#') {
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#6a9955' }}>{line.substring(i)}</span>);
                i = line.length;
                continue;
            }
            // Triple-quoted strings
            if ((line.substring(i, i + 3) === '"""' || line.substring(i, i + 3) === "'''")) {
                const quote = line.substring(i, i + 3);
                const end = line.indexOf(quote, i + 3);
                const strEnd = end !== -1 ? end + 3 : line.length;
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#ce9178' }}>{line.substring(i, strEnd)}</span>);
                i = strEnd;
                continue;
            }
            // Strings
            if (line[i] === '"' || line[i] === "'") {
                const quote = line[i];
                let j = i + 1;
                while (j < line.length && line[j] !== quote) {
                    if (line[j] === '\\') j++;
                    j++;
                }
                j = Math.min(j + 1, line.length);
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#ce9178' }}>{line.substring(i, j)}</span>);
                i = j;
                continue;
            }
            // Template references {{...}}
            if (line[i] === '{' && line[i + 1] === '{') {
                const end = line.indexOf('}}', i + 2);
                const refEnd = end !== -1 ? end + 2 : line.length;
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#4fc1ff', fontWeight: 'bold' }}>{line.substring(i, refEnd)}</span>);
                i = refEnd;
                continue;
            }
            // Decorators
            if (line[i] === '@') {
                let j = i + 1;
                while (j < line.length && /[\w.]/.test(line[j])) j++;
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#dcdcaa' }}>{line.substring(i, j)}</span>);
                i = j;
                continue;
            }
            // Numbers
            if (/\d/.test(line[i]) && (i === 0 || !/\w/.test(line[i - 1]))) {
                let j = i;
                while (j < line.length && /[\d.xXoObBeE_a-fA-F]/.test(line[j])) j++;
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#b5cea8' }}>{line.substring(i, j)}</span>);
                i = j;
                continue;
            }
            // Words (keywords, builtins, identifiers)
            if (/[a-zA-Z_]/.test(line[i])) {
                let j = i;
                while (j < line.length && /\w/.test(line[j])) j++;
                const word = line.substring(i, j);
                const kwSet = new Set(PYTHON_KEYWORDS.concat(['return', 'import', 'from', 'as', 'pass', 'break',
                    'continue', 'raise', 'yield', 'lambda', 'global', 'nonlocal', 'assert', 'del',
                    'in', 'not', 'and', 'or', 'is', 'True', 'False', 'None']));
                const builtins = new Set(['print', 'len', 'range', 'int', 'str', 'float', 'list', 'dict',
                    'set', 'tuple', 'bool', 'type', 'isinstance', 'enumerate', 'zip', 'map', 'filter',
                    'sorted', 'reversed', 'sum', 'min', 'max', 'abs', 'round', 'open', 'input',
                    'super', 'property', 'staticmethod', 'classmethod', 'hasattr', 'getattr', 'setattr']);
                if (kwSet.has(word)) {
                    tokens.push(<span key={`${li}-${i}`} style={{ color: '#569cd6', fontWeight: 'bold' }}>{word}</span>);
                } else if (builtins.has(word)) {
                    tokens.push(<span key={`${li}-${i}`} style={{ color: '#dcdcaa' }}>{word}</span>);
                } else if (word === 'self') {
                    tokens.push(<span key={`${li}-${i}`} style={{ color: '#9cdcfe' }}>{word}</span>);
                } else {
                    tokens.push(<span key={`${li}-${i}`} style={{ color: '#d4d4d4' }}>{word}</span>);
                }
                i = j;
                continue;
            }
            // Operators & punctuation
            if ('+-*/%=<>!&|^~:'.includes(line[i])) {
                tokens.push(<span key={`${li}-${i}`} style={{ color: '#d4d4d4' }}>{line[i]}</span>);
                i++;
                continue;
            }
            // Default
            tokens.push(<span key={`${li}-${i}`} style={{ color: '#d4d4d4' }}>{line[i]}</span>);
            i++;
        }
        return { lineNum: li + 1, tokens };
    });
}

/**
 * ScriptEditorPanel — a floating, draggable Python script editor panel
 * modeled after the MarkdownEditorPanel.
 *
 * Props:
 *  - cellRow, cellCol: cell coordinates
 *  - cellName: optional cell name (for {{name}} badge)
 *  - scriptText, setScriptText: controlled text state
 *  - scriptRowSpan, setScriptRowSpan: output row span
 *  - scriptColSpan, setScriptColSpan: output col span
 *  - canEdit: boolean
 *  - isOwner: whether current user is the owner (for Apply)
 *  - isLocked: whether cell is locked
 *  - onApply: () => void — called when Apply Script is clicked
 *  - onClose: () => void — called when Cancel / X is clicked
 *  - onInsertRange: () => void — called when Insert Range is clicked
 *  - textareaRef: ref to pass to the internal textarea (for insertSelectedRangeIntoScript)
 */
export default function ScriptEditorPanel({
    cellRow,
    cellCol,
    cellName,
    projectName,
    sheetName,
    scriptText,
    setScriptText,
    scriptRowSpan,
    setScriptRowSpan,
    scriptColSpan,
    setScriptColSpan,
    canEdit,
    isOwner,
    isLocked,
    onApply,
    onClose,
    onInsertRange,
    textareaRef: externalTextareaRef,
}) {
    const panelRef = useRef(null);
    const internalTextareaRef = useRef(null);
    const textareaRef = externalTextareaRef || internalTextareaRef;
    const highlightRef = useRef(null);

    const [isDragging, setIsDragging] = useState(false);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [position, setPosition] = useState({ x: null, y: null });
    const [isMaximized, setIsMaximized] = useState(false);
    const [activeTab, setActiveTab] = useState('edit'); // 'edit' | 'highlight' | 'preview'
    const [showLineNumbers, setShowLineNumbers] = useState(true);
    const [previewText, setPreviewText] = useState('');
    const [previewLoading, setPreviewLoading] = useState(false);
    const [previewError, setPreviewError] = useState('');

    // Fetch resolved preview from backend when the Preview tab is activated
    useEffect(() => {
        if (activeTab !== 'preview') return;
        setPreviewLoading(true);
        setPreviewError('');
        authenticatedFetch(
            apiUrl(`/api/preview/script?project=${encodeURIComponent(projectName || '')}&sheet=${encodeURIComponent(sheetName || '')}&row=${encodeURIComponent(String(cellRow))}&col=${encodeURIComponent(String(cellCol))}`),
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ script: scriptText || '' }),
            }
        )
            .then(r => r.ok ? r.json() : r.text().then(t => Promise.reject(t)))
            .then(data => { setPreviewText(data.resolved ?? ''); })
            .catch(err => { setPreviewError(String(err)); })
            .finally(() => setPreviewLoading(false));
    }, [activeTab, scriptText, projectName, sheetName, cellRow, cellCol]);

    // Position panel on mount
    useEffect(() => {
        if (position.x === null) {
            const vw = window.innerWidth;
            const vh = window.innerHeight;
            const panelWidth = 680;
            setPosition({
                x: Math.max(16, (vw - panelWidth) / 2),
                y: Math.max(16, vh * 0.1),
            });
        }
    }, []);

    // Dragging logic
    const handleMouseDown = useCallback((e) => {
        if (e.target.closest('.script-header')) {
            setIsDragging(true);
            const rect = panelRef.current.getBoundingClientRect();
            setDragOffset({ x: e.clientX - rect.left, y: e.clientY - rect.top });
        }
    }, []);

    useEffect(() => {
        if (!isDragging) return;
        const handleMouseMove = (e) => {
            setPosition({
                x: Math.max(0, e.clientX - dragOffset.x),
                y: Math.max(0, e.clientY - dragOffset.y),
            });
        };
        const handleMouseUp = () => setIsDragging(false);
        window.addEventListener('mousemove', handleMouseMove);
        window.addEventListener('mouseup', handleMouseUp);
        return () => {
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseup', handleMouseUp);
        };
    }, [isDragging, dragOffset]);

    // Sync scroll between textarea and highlight overlay
    const syncScroll = useCallback(() => {
        if (textareaRef.current && highlightRef.current) {
            highlightRef.current.scrollTop = textareaRef.current.scrollTop;
            highlightRef.current.scrollLeft = textareaRef.current.scrollLeft;
        }
    }, [textareaRef]);

    /**
     * Handle keydown in the textarea for Python-specific behavior:
     * - Tab inserts 4 spaces (or indents selection)
     * - Shift+Tab dedents
     * - Enter auto-indents, adds extra indent after ':'
     * - Backspace removes up to 4 trailing spaces (smart dedent)
     */
    const handleKeyDown = (e) => {
        const textarea = textareaRef.current;
        if (!textarea) return;
        const { selectionStart, selectionEnd } = textarea;

        // ── Tab: insert 4 spaces or indent selection ──
        if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey) {
            e.preventDefault();
            if (selectionStart === selectionEnd && !e.shiftKey) {
                // No selection — insert 4 spaces
                const before = scriptText.substring(0, selectionStart);
                const after = scriptText.substring(selectionEnd);
                const newText = before + '    ' + after;
                setScriptText(newText);
                setTimeout(() => {
                    textarea.selectionStart = textarea.selectionEnd = selectionStart + 4;
                }, 0);
            } else {
                // Selection exists — indent/dedent each line
                const textBefore = scriptText.substring(0, selectionStart);
                const lineStartIdx = textBefore.lastIndexOf('\n') + 1;
                const textAfter = scriptText.substring(selectionEnd);
                const selectedBlock = scriptText.substring(lineStartIdx, selectionEnd);
                const lines = selectedBlock.split('\n');
                let newLines;
                if (e.shiftKey) {
                    // Dedent: remove up to 4 leading spaces
                    newLines = lines.map(l => l.replace(/^ {1,4}/, ''));
                } else {
                    // Indent: add 4 spaces
                    newLines = lines.map(l => '    ' + l);
                }
                const newBlock = newLines.join('\n');
                const newText = scriptText.substring(0, lineStartIdx) + newBlock + textAfter;
                setScriptText(newText);
                setTimeout(() => {
                    textarea.selectionStart = lineStartIdx;
                    textarea.selectionEnd = lineStartIdx + newBlock.length;
                }, 0);
            }
            return;
        }

        // ── Shift+Tab without selection — dedent current line ──
        if (e.key === 'Tab' && e.shiftKey && selectionStart === selectionEnd) {
            e.preventDefault();
            const textBefore = scriptText.substring(0, selectionStart);
            const lineStartIdx = textBefore.lastIndexOf('\n') + 1;
            const lineEndIdx = scriptText.indexOf('\n', selectionStart);
            const actualEnd = lineEndIdx === -1 ? scriptText.length : lineEndIdx;
            const line = scriptText.substring(lineStartIdx, actualEnd);
            const dedented = line.replace(/^ {1,4}/, '');
            const removed = line.length - dedented.length;
            const newText = scriptText.substring(0, lineStartIdx) + dedented + scriptText.substring(actualEnd);
            setScriptText(newText);
            setTimeout(() => {
                const newPos = Math.max(lineStartIdx, selectionStart - removed);
                textarea.selectionStart = textarea.selectionEnd = newPos;
            }, 0);
            return;
        }

        // ── Enter: auto-indent ──
        if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
            e.preventDefault();
            const textBefore = scriptText.substring(0, selectionStart);
            const lineStartIdx = textBefore.lastIndexOf('\n') + 1;
            const currentLine = textBefore.substring(lineStartIdx);

            // Determine current indentation
            const indentMatch = currentLine.match(/^(\s*)/);
            let indent = indentMatch ? indentMatch[1] : '';

            // Check if line ends with ':' (after stripping comments and whitespace)
            const trimmedLine = currentLine.replace(/#.*$/, '').trimEnd();
            if (trimmedLine.endsWith(':')) {
                indent += '    ';
            }

            // Check for dedent keywords: return, break, continue, pass, raise
            const dedentKeywords = ['return', 'break', 'continue', 'pass', 'raise'];
            const lineContent = currentLine.trim();
            const firstWord = lineContent.split(/\s/)[0];
            if (dedentKeywords.includes(firstWord) && !trimmedLine.endsWith(':')) {
                // Optionally reduce indent for next line after a return/pass/etc.
                // Keep same indent (don't increase, but don't decrease — user can Shift+Tab)
            }

            const insertion = '\n' + indent;
            const newText = scriptText.substring(0, selectionStart) + insertion + scriptText.substring(selectionEnd);
            setScriptText(newText);
            setTimeout(() => {
                const newPos = selectionStart + insertion.length;
                textarea.selectionStart = textarea.selectionEnd = newPos;
                textarea.focus();
            }, 0);
            return;
        }

        // ── Backspace: smart dedent (remove up to 4 trailing spaces) ──
        if (e.key === 'Backspace' && selectionStart === selectionEnd && selectionStart > 0) {
            const textBefore = scriptText.substring(0, selectionStart);
            const lineStartIdx = textBefore.lastIndexOf('\n') + 1;
            const lineBeforeCursor = textBefore.substring(lineStartIdx);
            // If cursor is only preceded by spaces on this line
            if (/^\s+$/.test(lineBeforeCursor)) {
                const spacesToRemove = lineBeforeCursor.length % 4 === 0 ? 4 : lineBeforeCursor.length % 4;
                if (spacesToRemove > 0 && lineBeforeCursor.length >= spacesToRemove) {
                    e.preventDefault();
                    const newText = scriptText.substring(0, selectionStart - spacesToRemove) + scriptText.substring(selectionStart);
                    setScriptText(newText);
                    setTimeout(() => {
                        const newPos = selectionStart - spacesToRemove;
                        textarea.selectionStart = textarea.selectionEnd = newPos;
                    }, 0);
                    return;
                }
            }
        }

        // ── Auto-close brackets and quotes ──
        const pairs = { '(': ')', '[': ']', '{': '}', '"': '"', "'": "'" };
        if (pairs[e.key] && selectionStart === selectionEnd) {
            // Don't auto-close quotes if the previous char is alphanumeric (likely part of a word)
            if ((e.key === '"' || e.key === "'") && selectionStart > 0 && /\w/.test(scriptText[selectionStart - 1])) {
                return; // let default behavior
            }
            e.preventDefault();
            const before = scriptText.substring(0, selectionStart);
            const after = scriptText.substring(selectionEnd);
            const newText = before + e.key + pairs[e.key] + after;
            setScriptText(newText);
            setTimeout(() => {
                textarea.selectionStart = textarea.selectionEnd = selectionStart + 1;
            }, 0);
            return;
        }

        // ── Skip over closing bracket/quote if typed and matches next char ──
        const closers = new Set([')', ']', '}', '"', "'"]);
        if (closers.has(e.key) && selectionStart === selectionEnd && scriptText[selectionStart] === e.key) {
            e.preventDefault();
            setTimeout(() => {
                textarea.selectionStart = textarea.selectionEnd = selectionStart + 1;
            }, 0);
            return;
        }

        // ── Ctrl+/ : toggle comment on current line ──
        if ((e.ctrlKey || e.metaKey) && e.key === '/') {
            e.preventDefault();
            const textBefore = scriptText.substring(0, selectionStart);
            const lineStartIdx = textBefore.lastIndexOf('\n') + 1;
            const lineEndIdx = scriptText.indexOf('\n', selectionStart);
            const actualEnd = lineEndIdx === -1 ? scriptText.length : lineEndIdx;
            const line = scriptText.substring(lineStartIdx, actualEnd);
            const indentMatch = line.match(/^(\s*)/);
            const indent = indentMatch ? indentMatch[1] : '';
            const content = line.substring(indent.length);
            let newLine;
            let cursorDelta;
            if (content.startsWith('# ')) {
                newLine = indent + content.substring(2);
                cursorDelta = -2;
            } else if (content.startsWith('#')) {
                newLine = indent + content.substring(1);
                cursorDelta = -1;
            } else {
                newLine = indent + '# ' + content;
                cursorDelta = 2;
            }
            const newText = scriptText.substring(0, lineStartIdx) + newLine + scriptText.substring(actualEnd);
            setScriptText(newText);
            setTimeout(() => {
                const newPos = Math.max(lineStartIdx, selectionStart + cursorDelta);
                textarea.selectionStart = textarea.selectionEnd = newPos;
            }, 0);
            return;
        }
    };

    const lineCount = (scriptText || '').split('\n').length;
    const highlighted = highlightPython(scriptText || '');

    const panelWidth = isMaximized ? '85vw' : '680px';
    const panelHeight = isMaximized ? '80vh' : '520px';

    return (
        <div
            ref={panelRef}
            className="shadow-lg"
            style={{
                position: 'fixed',
                left: isMaximized ? '7.5vw' : (position.x ?? 'auto'),
                top: isMaximized ? '10vh' : (position.y ?? 80),
                width: panelWidth,
                height: panelHeight,
                zIndex: 2100,
                display: 'flex',
                flexDirection: 'column',
                background: '#1e1e1e',
                border: '1px solid #333',
                borderRadius: '12px',
                overflow: 'hidden',
                resize: isMaximized ? 'none' : 'both',
                minWidth: '420px',
                minHeight: '320px',
            }}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
        >
            {/* Header — draggable */}
            <div
                className="script-header d-flex align-items-center justify-content-between px-3 py-2"
                style={{
                    background: 'linear-gradient(135deg, #1a7a4c 0%, #0d5a3e 100%)',
                    color: 'white',
                    cursor: isDragging ? 'grabbing' : 'grab',
                    userSelect: 'none',
                    flexShrink: 0,
                }}
                onMouseDown={handleMouseDown}
            >
                <div className="d-flex align-items-center gap-2">
                    <Code size={16} />
                    <span className="fw-semibold small">
                        Python Script — Cell {cellCol}{cellRow}
                    </span>
                    {cellName && (
                        <span
                            className="px-1 rounded"
                            style={{
                                background: 'rgba(255,255,255,0.2)',
                                color: '#a5f3fc',
                                fontFamily: 'monospace',
                                fontSize: '0.7rem',
                            }}
                            title={`Cell name — use {{${cellName}}} to reference this cell`}
                        >
                            {`{{${cellName}}}`}
                        </span>
                    )}
                </div>
                <div className="d-flex align-items-center gap-1">
                    <button
                        className="btn btn-sm p-1"
                        style={{ color: 'white', background: 'transparent', border: 'none' }}
                        onClick={() => setIsMaximized(v => !v)}
                        title={isMaximized ? 'Restore' : 'Maximize'}
                    >
                        {isMaximized ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
                    </button>
                    <button
                        className="btn btn-sm p-1"
                        style={{ color: 'white', background: 'transparent', border: 'none' }}
                        onClick={onClose}
                        title="Close"
                    >
                        <X size={16} />
                    </button>
                </div>
            </div>

            {/* Tabs */}
            <div className="d-flex border-bottom" style={{ flexShrink: 0, background: '#252526', borderColor: '#333 !important' }}>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0`}
                    style={{
                        borderBottom: activeTab === 'edit' ? '2px solid #1a7a4c' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'edit' ? '#4ec9b0' : '#888',
                        fontWeight: activeTab === 'edit' ? 'bold' : 'normal',
                    }}
                    onClick={() => setActiveTab('edit')}
                >
                    <Code size={12} className="me-1" /> Edit
                </button>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0`}
                    style={{
                        borderBottom: activeTab === 'highlight' ? '2px solid #1a7a4c' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'highlight' ? '#4ec9b0' : '#888',
                        fontWeight: activeTab === 'highlight' ? 'bold' : 'normal',
                    }}
                    onClick={() => setActiveTab('highlight')}
                >
                    <Code size={12} className="me-1" /> Highlighted
                </button>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0`}
                    style={{
                        borderBottom: activeTab === 'preview' ? '2px solid #1a7a4c' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'preview' ? '#4ec9b0' : '#888',
                        fontWeight: activeTab === 'preview' ? 'bold' : 'normal',
                    }}
                    onClick={() => setActiveTab('preview')}
                >
                    <Eye size={12} className="me-1" /> Preview
                </button>
            </div>

            {/* Editor body */}
            <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                {/* Settings bar */}
                <div className="d-flex align-items-center gap-3 px-3 py-1" style={{ background: '#2d2d2d', borderBottom: '1px solid #333', flexShrink: 0 }}>
                    <div className="d-flex align-items-center gap-1">
                        <label className="small" style={{ color: '#aaa', whiteSpace: 'nowrap', fontSize: '0.75rem' }}>Output Rows</label>
                        <input
                            type="number"
                            min={1}
                            className="form-control form-control-sm"
                            style={{ width: 56, background: '#3c3c3c', color: '#d4d4d4', border: '1px solid #555', fontSize: '0.75rem' }}
                            value={scriptRowSpan}
                            onChange={(e) => setScriptRowSpan(Math.max(1, parseInt(e.target.value, 10) || 1))}
                            disabled={!canEdit}
                            title="Script output row span"
                        />
                    </div>
                    <div className="d-flex align-items-center gap-1">
                        <label className="small" style={{ color: '#aaa', whiteSpace: 'nowrap', fontSize: '0.75rem' }}>Output Cols</label>
                        <input
                            type="number"
                            min={1}
                            className="form-control form-control-sm"
                            style={{ width: 56, background: '#3c3c3c', color: '#d4d4d4', border: '1px solid #555', fontSize: '0.75rem' }}
                            value={scriptColSpan}
                            onChange={(e) => setScriptColSpan(Math.max(1, parseInt(e.target.value, 10) || 1))}
                            disabled={!canEdit}
                            title="Script output column span"
                        />
                    </div>
                    <label className="d-flex align-items-center gap-1 small" style={{ color: '#aaa', cursor: 'pointer', fontSize: '0.75rem' }}>
                        <input
                            type="checkbox"
                            checked={showLineNumbers}
                            onChange={(e) => setShowLineNumbers(e.target.checked)}
                            style={{ accentColor: '#1a7a4c' }}
                        />
                        Line #
                    </label>
                    <div className="d-flex align-items-center gap-1 ms-auto" style={{ fontSize: '0.7rem', color: '#666' }}>
                        <span title="Tab inserts 4 spaces. Shift+Tab dedents. Ctrl+/ toggles comment.">
                            Tab=4sp | Ctrl+/ comment | Auto-indent
                        </span>
                    </div>
                </div>

                {/* Code area */}
                <div style={{ flex: 1, position: 'relative', overflow: 'hidden' }}>
                    {activeTab === 'edit' && (
                        <div style={{ position: 'relative', width: '100%', height: '100%', display: 'flex' }}>
                            {/* Line numbers gutter */}
                            {showLineNumbers && (
                                <div
                                    style={{
                                        width: 44,
                                        flexShrink: 0,
                                        background: '#1e1e1e',
                                        borderRight: '1px solid #333',
                                        overflowY: 'hidden',
                                        fontFamily: "'Consolas', 'Courier New', 'Monaco', monospace",
                                        fontSize: '13px',
                                        lineHeight: '20px',
                                        paddingTop: 8,
                                        color: '#858585',
                                        textAlign: 'right',
                                        paddingRight: 8,
                                        userSelect: 'none',
                                        pointerEvents: 'none',
                                    }}
                                    ref={(el) => {
                                        // Sync line number scroll with textarea
                                        if (el && textareaRef.current) {
                                            const ta = textareaRef.current;
                                            const handler = () => { el.scrollTop = ta.scrollTop; };
                                            ta.removeEventListener('scroll', handler);
                                            ta.addEventListener('scroll', handler);
                                        }
                                    }}
                                    id="script-line-numbers"
                                >
                                    {Array.from({ length: lineCount }, (_, i) => (
                                        <div key={i + 1} style={{ height: 20 }}>{i + 1}</div>
                                    ))}
                                </div>
                            )}
                            <textarea
                                ref={textareaRef}
                                style={{
                                    flex: 1,
                                    height: '100%',
                                    background: '#1e1e1e',
                                    color: '#d4d4d4',
                                    border: 'none',
                                    outline: 'none',
                                    resize: 'none',
                                    fontFamily: "'Consolas', 'Courier New', 'Monaco', monospace",
                                    fontSize: '13px',
                                    lineHeight: '20px',
                                    padding: '8px 12px',
                                    tabSize: 4,
                                    whiteSpace: 'pre',
                                    overflowWrap: 'normal',
                                    overflowX: 'auto',
                                    caretColor: '#aeafad',
                                }}
                                spellCheck={false}
                                value={scriptText}
                                onChange={(e) => setScriptText(e.target.value)}
                                onKeyDown={handleKeyDown}
                                onScroll={syncScroll}
                                disabled={!canEdit}
                                placeholder={`# Python script for cell ${cellCol}${cellRow}\n# Use {{CellRef}} to reference other cells\n# Press Tab for indent, Shift+Tab to dedent\n# Ctrl+/ to toggle comment`}
                            />
                        </div>
                    )}

                    {activeTab === 'highlight' && (
                        <div
                            style={{
                                width: '100%',
                                height: '100%',
                                overflowY: 'auto',
                                overflowX: 'auto',
                                background: '#1e1e1e',
                                padding: '8px 0',
                                fontFamily: "'Consolas', 'Courier New', 'Monaco', monospace",
                                fontSize: '13px',
                                lineHeight: '20px',
                            }}
                        >
                            {highlighted.map(({ lineNum, tokens }) => (
                                <div key={lineNum} style={{ display: 'flex', minHeight: 20 }}>
                                    {showLineNumbers && (
                                        <span style={{
                                            width: 44,
                                            flexShrink: 0,
                                            textAlign: 'right',
                                            paddingRight: 8,
                                            color: '#858585',
                                            userSelect: 'none',
                                            borderRight: '1px solid #333',
                                            marginRight: 12,
                                        }}>
                                            {lineNum}
                                        </span>
                                    )}
                                    <span style={{ whiteSpace: 'pre' }}>
                                        {tokens.length > 0 ? tokens : ' '}
                                    </span>
                                </div>
                            ))}
                        </div>
                    )}

                    {activeTab === 'preview' && (
                        <div
                            style={{
                                width: '100%',
                                height: '100%',
                                overflowY: 'auto',
                                overflowX: 'auto',
                                background: '#1e1e1e',
                                padding: '8px 0',
                                fontFamily: "'Consolas', 'Courier New', 'Monaco', monospace",
                                fontSize: '13px',
                                lineHeight: '20px',
                            }}
                        >
                            {previewLoading && (
                                <div style={{ padding: '8px 16px', color: '#888', fontStyle: 'italic' }}>Resolving references…</div>
                            )}
                            {!previewLoading && previewError && (
                                <div style={{ padding: '8px 16px', color: '#f48771' }}>Error: {previewError}</div>
                            )}
                            {!previewLoading && !previewError && (
                                previewText
                                    ? highlightPython(previewText).map(({ lineNum, tokens }) => (
                                        <div key={lineNum} style={{ display: 'flex', minHeight: 20 }}>
                                            {showLineNumbers && (
                                                <span style={{
                                                    width: 44,
                                                    flexShrink: 0,
                                                    textAlign: 'right',
                                                    paddingRight: 8,
                                                    color: '#858585',
                                                    userSelect: 'none',
                                                    borderRight: '1px solid #333',
                                                    marginRight: 12,
                                                }}>
                                                    {lineNum}
                                                </span>
                                            )}
                                            <span style={{ whiteSpace: 'pre' }}>
                                                {tokens.length > 0 ? tokens : ' '}
                                            </span>
                                        </div>
                                    ))
                                    : <div style={{ padding: '8px 16px', color: '#555', fontStyle: 'italic' }}>Nothing to preview.</div>
                            )}
                        </div>
                    )}
                </div>
            </div>

            {/* Footer — actions */}
            <div
                className="d-flex align-items-center justify-content-between px-3 py-2"
                style={{
                    background: '#252526',
                    borderTop: '1px solid #333',
                    flexShrink: 0,
                }}
            >
                <div className="d-flex align-items-center gap-2">
                    <button
                        className="btn btn-sm"
                        style={{
                            background: '#2ea04e',
                            color: 'white',
                            border: 'none',
                            fontSize: '0.8rem',
                        }}
                        onClick={onInsertRange}
                        disabled={!canEdit}
                        title="Insert selected range into script at cursor position"
                    >
                        <CornerDownLeft size={13} className="me-1" />
                        Insert Range
                    </button>
                    <span style={{ color: '#666', fontSize: '0.7rem' }}>
                        {lineCount} line{lineCount !== 1 ? 's' : ''} · {(scriptText || '').length} chars
                    </span>
                </div>
                <div className="d-flex align-items-center gap-2">
                    <button
                        className="btn btn-sm"
                        style={{
                            background: '#0e639c',
                            color: 'white',
                            border: 'none',
                            fontSize: '0.8rem',
                        }}
                        onClick={onApply}
                        disabled={!canEdit || !isOwner || isLocked}
                        title="Apply script (owner only)"
                    >
                        <Play size={13} className="me-1" />
                        Apply Script
                    </button>
                    <button
                        className="btn btn-sm"
                        style={{
                            background: '#3c3c3c',
                            color: '#ccc',
                            border: '1px solid #555',
                            fontSize: '0.8rem',
                        }}
                        onClick={onClose}
                        title="Cancel"
                    >
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );
}
