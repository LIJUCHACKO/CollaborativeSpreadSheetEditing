import React, { useState, useRef, useEffect, useCallback } from 'react';
import { marked } from 'marked';
import { X, Eye, Edit3, Bold, Italic, Heading1, Heading2, Heading3, List, ListOrdered, Code, Link, Image, Quote, Minus, CheckSquare, Maximize2, Minimize2, Sigma, Table, FolderOpen, FileCode } from 'lucide-react';
import AssetBrowser from './AssetBrowser';
import PythonFileBrowser from './PythonFileBrowser';

// Configure marked for safe rendering
marked.setOptions({
    breaks: true,
    gfm: true,
});

// Load MathJax from npm once per page lifetime
function ensureMathJax() {
    if (typeof window === 'undefined') return;
    if (window.MathJax && window.MathJax.typesetPromise) return; // already loaded
    // Configure MathJax before loading
    window.MathJax = {
        tex: {
            inlineMath: [['$', '$'], ['\\(', '\\)']],
            displayMath: [['$$', '$$'], ['\\[', '\\]']],
            processEscapes: true,
            packages: { '[+]': ['ams', 'boldsymbol'] },
        },
        options: {
            skipHtmlTags: ['script', 'noscript', 'style', 'textarea', 'pre'],
        },
        startup: { typeset: false },
    };
    import('mathjax/es5/tex-chtml').catch(() => {});
}

export default function MarkdownEditorPanel({ cellRow, cellCol, value, onSave, onClose, readOnly, project }) {
    const [content, setContent] = useState(value || '');
    const [activeTab, setActiveTab] = useState('edit'); // 'edit' | 'preview' | 'split'
    const [isMaximized, setIsMaximized] = useState(false);
    const [assetBrowserOpen, setAssetBrowserOpen] = useState(false);
    const [pythonFileBrowserOpen, setPythonFileBrowserOpen] = useState(false);
    const textareaRef = useRef(null);
    const panelRef = useRef(null);
    const previewRef = useRef(null);
    const [isDragging, setIsDragging] = useState(false);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [position, setPosition] = useState({ x: null, y: null });
    // Table insert popover
    const [tablePopover, setTablePopover] = useState(false);
    const [tableRows, setTableRows] = useState(3);
    const [tableCols, setTableCols] = useState(3);
    // Table context menu
    const [tableMenu, setTableMenu] = useState(null); // { x, y, cursorPos }
    // Always holds the latest content so effects/cleanup can access it without stale closures
    const contentRef = useRef(content);
    useEffect(() => { contentRef.current = content; }, [content]);

    // Ensure MathJax CDN is loaded
    useEffect(() => { ensureMathJax(); }, []);

    // Auto-save previous cell's content when the target cell changes
    useEffect(() => {
        // On every cell change (after the first mount), save the content that was
        // being edited for the *previous* cell before switching to the new one.
        return () => {
            if (!readOnly) {
                onSave(contentRef.current);
            }
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [cellRow, cellCol]);

    // Sync incoming value when cell changes
    useEffect(() => {
        setContent(value || '');
    }, [value, cellRow, cellCol]);

    // Position panel on mount
    useEffect(() => {
        if (position.x === null) {
            const vw = window.innerWidth;
            const panelWidth = isMaximized ? vw * 0.8 : 720;
            setPosition({
                x: Math.max(16, vw - panelWidth - 32),
                y: 80,
            });
        }
    }, []);

    // Run MathJax typesetting whenever preview is visible and content changes
    useEffect(() => {
        if (activeTab !== 'preview' && activeTab !== 'split') return;
        if (!previewRef.current) return;
        const typesetPreview = () => {
            if (window.MathJax && window.MathJax.typesetPromise) {
                window.MathJax.typesetPromise([previewRef.current]).catch(() => {});
            }
        };
        // If MathJax is already ready, typeset immediately; else wait for it
        if (window.MathJax && window.MathJax.typesetPromise) {
            typesetPreview();
        } else {
            // Poll until MathJax is ready (CDN async load)
            const interval = setInterval(() => {
                if (window.MathJax && window.MathJax.typesetPromise) {
                    clearInterval(interval);
                    typesetPreview();
                }
            }, 200);
            return () => clearInterval(interval);
        }
    }, [content, activeTab]);

    // Dragging logic
    const handleMouseDown = useCallback((e) => {
        if (e.target.closest('.md-toolbar') || e.target.closest('.md-header')) {
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

    // Insert markdown syntax at cursor
    const insertMarkdown = (prefix, suffix = '', placeholder = '') => {
        const textarea = textareaRef.current;
        if (!textarea) return;
        const start = textarea.selectionStart;
        const end = textarea.selectionEnd;
        const selected = content.substring(start, end);
        const text = selected || placeholder;
        const newContent = content.substring(0, start) + prefix + text + suffix + content.substring(end);
        setContent(newContent);
        // Restore cursor
        setTimeout(() => {
            textarea.focus();
            const cursorPos = start + prefix.length + text.length;
            textarea.setSelectionRange(
                selected ? start + prefix.length : cursorPos,
                cursorPos
            );
        }, 0);
    };

    // ── Table helpers ────────────────────────────────────────────────────────

    /** Build a blank GFM markdown table string */
    const buildTableMarkdown = (rows, cols) => {
        const header = '| ' + Array.from({ length: cols }, (_, i) => `Header ${i + 1}`).join(' | ') + ' |';
        const divider = '| ' + Array(cols).fill('---').join(' | ') + ' |';
        const dataRow = '| ' + Array(cols).fill('   ').join(' | ') + ' |';
        return [header, divider, ...Array(rows - 1).fill(dataRow)].join('\n');
    };

    /** Insert a blank table at cursor */
    const insertTable = () => {
        const r = Math.max(2, tableRows);
        const c = Math.max(1, tableCols);
        const tableStr = '\n' + buildTableMarkdown(r, c) + '\n';
        insertMarkdown(tableStr);
        setTablePopover(false);
    };

    /**
     * Given the full content string and a cursor position inside it,
     * return info about the markdown table the cursor is in, or null.
     * Returns { tableStart, tableEnd, rows: string[][], rawLines: string[], lineIndex }
     * where rows[0] is the header, rows[1] is the divider row (kept as-is), rows[2..] are data.
     */
    const getTableAtCursor = (text, cursorPos) => {
        const lines = text.split('\n');
        // Find which line the cursor is on
        let charCount = 0;
        let cursorLine = -1;
        const lineStarts = [];
        for (let i = 0; i < lines.length; i++) {
            lineStarts.push(charCount);
            if (cursorLine === -1 && charCount + lines[i].length >= cursorPos) cursorLine = i;
            charCount += lines[i].length + 1;
        }
        if (cursorLine === -1) cursorLine = lines.length - 1;

        const isTableLine = (l) => l.trim().startsWith('|') && l.trim().endsWith('|');
        if (!isTableLine(lines[cursorLine])) return null;

        // Expand up
        let start = cursorLine;
        while (start > 0 && isTableLine(lines[start - 1])) start--;
        // Expand down
        let end = cursorLine;
        while (end < lines.length - 1 && isTableLine(lines[end + 1])) end++;

        if (end - start < 1) return null; // need at least header + divider

        const rawLines = lines.slice(start, end + 1);
        const tableStart = lineStarts[start];
        const tableEnd = lineStarts[end] + lines[end].length;

        // Parse each row into cells
        const parseRow = (line) =>
            line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());

        const rows = rawLines.map(parseRow);
        return { tableStart, tableEnd, rows, rawLines, lineIndex: cursorLine - start };
    };

    /** Rebuild a markdown table from a rows array (rows[1] stays as divider). */
    const serializeTable = (rows) => {
        // Compute column widths
        const cols = rows[0].length;
        const widths = Array.from({ length: cols }, (_, ci) =>
            Math.max(3, ...rows.map(r => (r[ci] || '').length))
        );
        return rows.map((row, ri) => {
            if (ri === 1) {
                // divider row
                return '| ' + Array.from({ length: cols }, (_, ci) => '-'.repeat(widths[ci])).join(' | ') + ' |';
            }
            return '| ' + Array.from({ length: cols }, (_, ci) => (row[ci] || '').padEnd(widths[ci])).join(' | ') + ' |';
        }).join('\n');
    };

    const applyTableEdit = (editFn) => {
        const textarea = textareaRef.current;
        const cursorPos = textarea ? textarea.selectionStart : content.length;
        const info = getTableAtCursor(content, cursorPos);
        if (!info) return;
        const { tableStart, tableEnd, rows } = info;
        const newRows = editFn(rows);
        if (!newRows) return;
        const newTable = serializeTable(newRows);
        const newContent = content.substring(0, tableStart) + newTable + content.substring(tableEnd);
        setContent(newContent);
        setTimeout(() => { if (textarea) textarea.focus(); }, 0);
    };

    const tableOps = {
        insertRowAbove: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const { rows, lineIndex } = info;
            // lineIndex 0 = header, 1 = divider — insert after divider at minimum
            const insertAt = Math.max(2, lineIndex);
            applyTableEdit(r => [
                ...r.slice(0, insertAt),
                Array(r[0].length).fill(''),
                ...r.slice(insertAt),
            ]);
        },
        insertRowBelow: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const { rows, lineIndex } = info;
            const insertAt = Math.max(2, lineIndex + 1);
            applyTableEdit(r => [
                ...r.slice(0, insertAt),
                Array(r[0].length).fill(''),
                ...r.slice(insertAt),
            ]);
        },
        removeRow: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const { rows, lineIndex } = info;
            if (lineIndex <= 1) return; // can't remove header/divider
            if (rows.length <= 3) return; // keep at least 1 data row
            applyTableEdit(r => r.filter((_, i) => i !== lineIndex));
        },
        insertColLeft: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const colIdx = getColAtCursor(info, cursorPos);
            applyTableEdit(r => r.map((row, ri) =>
                [...row.slice(0, colIdx), ri === 1 ? '---' : '', ...row.slice(colIdx)]
            ));
        },
        insertColRight: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const colIdx = getColAtCursor(info, cursorPos);
            applyTableEdit(r => r.map((row, ri) =>
                [...row.slice(0, colIdx + 1), ri === 1 ? '---' : '', ...row.slice(colIdx + 1)]
            ));
        },
        removeCol: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            if (info.rows[0].length <= 1) return;
            const colIdx = getColAtCursor(info, cursorPos);
            applyTableEdit(r => r.map(row => row.filter((_, ci) => ci !== colIdx)));
        },
        /** Pad all cells in the cursor's column to the same width (max content width). */
        prettifyCol: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            const colIdx = getColAtCursor(info, cursorPos);
            applyTableEdit(rows => {
                // Width = max of all non-divider cells in this column, minimum 3
                const maxW = Math.max(3, ...rows
                    .filter((_, ri) => ri !== 1)
                    .map(row => (row[colIdx] || '').length)
                );
                return rows.map((row, ri) => {
                    const newRow = [...row];
                    if (ri === 1) {
                        newRow[colIdx] = '-'.repeat(maxW);
                    } else {
                        newRow[colIdx] = (row[colIdx] || '').padEnd(maxW);
                    }
                    return newRow;
                });
            });
        },
        /** Pad every column so all cells share the same width — fully pretty table. */
        prettifyTable: (cursorPos) => {
            const info = getTableAtCursor(content, cursorPos);
            if (!info) return;
            // serializeTable already computes per-column max widths, so just re-serialize
            const { tableStart, tableEnd, rows } = info;
            const newTable = serializeTable(rows);
            const newContent = content.substring(0, tableStart) + newTable + content.substring(tableEnd);
            setContent(newContent);
            setTimeout(() => { if (textareaRef.current) textareaRef.current.focus(); }, 0);
        },
    };

    /** Determine which column the cursor is in within the table row */
    const getColAtCursor = (info, cursorPos) => {
        const lines = content.split('\n');
        let charCount = 0;
        const lineStarts = [];
        for (let i = 0; i < lines.length; i++) {
            lineStarts.push(charCount);
            charCount += lines[i].length + 1;
        }
        // find line index of tableStart
        let tableLineIdx = lineStarts.findIndex((s, i) => s === info.tableStart || (s <= info.tableStart && (lineStarts[i + 1] === undefined || lineStarts[i + 1] > info.tableStart)));
        const cursorLineIdx = tableLineIdx + info.lineIndex;
        const lineStart = lineStarts[cursorLineIdx] || info.tableStart;
        const posInLine = cursorPos - lineStart;
        const line = lines[cursorLineIdx] || '';
        // Count pipes before cursor position
        let pipeCount = 0;
        for (let i = 0; i < Math.min(posInLine, line.length); i++) {
            if (line[i] === '|') pipeCount++;
        }
        return Math.max(0, pipeCount - 1);
    };

    /** Handle right-click in textarea to show table context menu */
    const handleTextareaContextMenu = (e) => {
        const textarea = textareaRef.current;
        if (!textarea) return;
        const info = getTableAtCursor(content, textarea.selectionStart);
        if (!info) return;
        e.preventDefault();
        setTableMenu({ x: e.clientX, y: e.clientY, cursorPos: textarea.selectionStart });
    };

    // ── End table helpers ────────────────────────────────────────────────────

    const toolbarActions = [
        { icon: <Bold size={14} />, title: 'Bold (Ctrl+B)', action: () => insertMarkdown('**', '**', 'bold text') },
        { icon: <Italic size={14} />, title: 'Italic (Ctrl+I)', action: () => insertMarkdown('*', '*', 'italic text') },
        { type: 'separator' },
        { icon: <Heading1 size={14} />, title: 'Heading 1', action: () => insertMarkdown('# ', '', 'Heading') },
        { icon: <Heading2 size={14} />, title: 'Heading 2', action: () => insertMarkdown('## ', '', 'Heading') },
        { icon: <Heading3 size={14} />, title: 'Heading 3', action: () => insertMarkdown('### ', '', 'Heading') },
        { type: 'separator' },
        { icon: <List size={14} />, title: 'Unordered List', action: () => insertMarkdown('- ', '', 'list item') },
        { icon: <ListOrdered size={14} />, title: 'Ordered List', action: () => insertMarkdown('1. ', '', 'list item') },
        { icon: <CheckSquare size={14} />, title: 'Task List', action: () => insertMarkdown('- [ ] ', '', 'task') },
        { type: 'separator' },
        { icon: <Code size={14} />, title: 'Inline Code', action: () => insertMarkdown('`', '`', 'code') },
        { icon: <Quote size={14} />, title: 'Blockquote', action: () => insertMarkdown('> ', '', 'quote') },
        { icon: <Minus size={14} />, title: 'Horizontal Rule', action: () => insertMarkdown('\n---\n') },
        { type: 'separator' },
        { icon: <Link size={14} />, title: 'Link', action: () => insertMarkdown('[', '](url)', 'link text') },
        { icon: <Image size={14} />, title: 'Image', action: () => insertMarkdown('![', '](url)', 'alt text') },
        { type: 'separator' },
        { icon: <Sigma size={14} />, title: 'Inline Equation  $...$', action: () => insertMarkdown('$', '$', 'equation') },
        {
            icon: <span style={{ fontFamily: 'serif', fontStyle: 'italic', fontSize: '13px', lineHeight: 1 }}>∑²</span>,
            title: 'Display Equation  $$...$$',
            action: () => insertMarkdown('\n$$\n', '\n$$\n', 'equation'),
        },
        { type: 'separator' },
        { icon: <Table size={14} />, title: 'Insert Table', action: () => setTablePopover(v => !v), isTableBtn: true },
        { type: 'separator' },
        { icon: <FolderOpen size={14} />, title: 'Browse & Insert Asset Image', action: () => setAssetBrowserOpen(true) },
        { icon: <FileCode size={14} />, title: 'Browse & Insert Python File Link', action: () => setPythonFileBrowserOpen(true) },
    ];

    // Keyboard shortcuts
    const handleKeyDown = (e) => {
        if (e.ctrlKey || e.metaKey) {
            if (e.key === 'b') { e.preventDefault(); insertMarkdown('**', '**', 'bold text'); }
            if (e.key === 'i') { e.preventDefault(); insertMarkdown('*', '*', 'italic text'); }
            if (e.key === 's') { e.preventDefault(); onSave(content); }
        }

        // Continue list on Enter
        if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
            const textarea = textareaRef.current;
            if (!textarea) return;
            const cursorPos = textarea.selectionStart;
            const textBefore = content.substring(0, cursorPos);
            const currentLineStart = textBefore.lastIndexOf('\n') + 1;
            const currentLine = textBefore.substring(currentLineStart);

            // Match list prefixes:
            // - unordered:  "- ", "* ", "+ ", "  - " (nested), etc.
            // - task list:  "- [ ] ", "- [x] "
            // - ordered:    "1. ", "2. ", etc.
            const unorderedMatch = currentLine.match(/^(\s*)([-*+])\s(\[[ xX]\]\s)?/);
            const orderedMatch = currentLine.match(/^(\s*)(\d+)\.\s/);

            if (unorderedMatch) {
                const lineContent = currentLine.substring(unorderedMatch[0].length);
                if (lineContent.trim() === '') {
                    // Empty list item → break out of the list
                    e.preventDefault();
                    const indent = unorderedMatch[1];
                    const newContent =
                        content.substring(0, currentLineStart) +
                        '\n' +
                        content.substring(cursorPos);
                    setContent(newContent);
                    setTimeout(() => {
                        const newPos = currentLineStart + 1;
                        textarea.setSelectionRange(newPos, newPos);
                    }, 0);
                } else {
                    e.preventDefault();
                    const indent = unorderedMatch[1];
                    const bullet = unorderedMatch[2];
                    // For task lists, continue with an unchecked item
                    const taskPrefix = unorderedMatch[3] ? '[ ] ' : '';
                    const continuation = `\n${indent}${bullet} ${taskPrefix}`;
                    const newContent =
                        content.substring(0, cursorPos) +
                        continuation +
                        content.substring(textarea.selectionEnd);
                    setContent(newContent);
                    setTimeout(() => {
                        const newPos = cursorPos + continuation.length;
                        textarea.setSelectionRange(newPos, newPos);
                    }, 0);
                }
            } else if (orderedMatch) {
                const lineContent = currentLine.substring(orderedMatch[0].length);
                if (lineContent.trim() === '') {
                    // Empty ordered item → break out of the list
                    e.preventDefault();
                    const newContent =
                        content.substring(0, currentLineStart) +
                        '\n' +
                        content.substring(cursorPos);
                    setContent(newContent);
                    setTimeout(() => {
                        const newPos = currentLineStart + 1;
                        textarea.setSelectionRange(newPos, newPos);
                    }, 0);
                } else {
                    e.preventDefault();
                    const indent = orderedMatch[1];
                    const nextNum = parseInt(orderedMatch[2], 10) + 1;
                    const continuation = `\n${indent}${nextNum}. `;
                    const newContent =
                        content.substring(0, cursorPos) +
                        continuation +
                        content.substring(textarea.selectionEnd);
                    setContent(newContent);
                    setTimeout(() => {
                        const newPos = cursorPos + continuation.length;
                        textarea.setSelectionRange(newPos, newPos);
                    }, 0);
                }
            }
        }
    };

    const getHtml = () => {
        try {
            return { __html: marked.parse(content || '') };
        } catch {
            return { __html: '<p class="text-danger">Error rendering markdown</p>' };
        }
    };

    const panelWidth = isMaximized ? '80vw' : '620px';
    const panelHeight = isMaximized ? '80vh' : '780px';

    return (
        <div
            ref={panelRef}
            className="shadow-lg"
            style={{
                position: 'fixed',
                left: isMaximized ? '10vw' : position.x ?? 'auto',
                top: isMaximized ? '10vh' : position.y ?? 80,
                width: panelWidth,
                height: panelHeight,
                zIndex: 2000,
                display: 'flex',
                flexDirection: 'column',
                background: '#ffffff',
                border: '1px solid #d1d5db',
                borderRadius: '12px',
                overflow: 'hidden',
                resize: isMaximized ? 'none' : 'both',
                minWidth: '380px',
                minHeight: '480px',
            }}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
        >
            {/* Header */}
            <div
                className="md-header d-flex align-items-center justify-content-between px-3 py-2"
                style={{
                    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                    color: 'white',
                    cursor: isDragging ? 'grabbing' : 'grab',
                    userSelect: 'none',
                    flexShrink: 0,
                }}
                onMouseDown={handleMouseDown}
            >
                <div className="d-flex align-items-center gap-2">
                    <Edit3 size={16} />
                    <span className="fw-semibold small">
                        Markdown Editor — Cell {cellCol}{cellRow}
                    </span>
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
                        onClick={() => { if (!readOnly) onSave(contentRef.current); onClose(); }}
                        title="Close"
                    >
                        <X size={16} />
                    </button>
                </div>
            </div>

            {/* Tabs */}
            <div className="d-flex border-bottom" style={{ flexShrink: 0, background: '#f9fafb' }}>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0 ${activeTab === 'edit' ? 'fw-bold' : 'text-muted'}`}
                    style={{
                        borderBottom: activeTab === 'edit' ? '2px solid #667eea' : '2px solid transparent',
                        background: 'transparent',
                    }}
                    onClick={() => setActiveTab('edit')}
                >
                    <Edit3 size={12} className="me-1" /> Edit
                </button>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0 ${activeTab === 'preview' ? 'fw-bold' : 'text-muted'}`}
                    style={{
                        borderBottom: activeTab === 'preview' ? '2px solid #667eea' : '2px solid transparent',
                        background: 'transparent',
                    }}
                    onClick={() => setActiveTab('preview')}
                >
                    <Eye size={12} className="me-1" /> Preview
                </button>
                <button
                    className={`btn btn-sm px-3 py-1 rounded-0 border-0 ${activeTab === 'split' ? 'fw-bold' : 'text-muted'}`}
                    style={{
                        borderBottom: activeTab === 'split' ? '2px solid #667eea' : '2px solid transparent',
                        background: 'transparent',
                    }}
                    onClick={() => setActiveTab('split')}
                >
                    <Eye size={12} className="me-1" /> Split
                </button>
            </div>

            {/* Toolbar (visible in edit & split mode) */}
            {(activeTab === 'edit' || activeTab === 'split') && !readOnly && (
                <div
                    className="md-toolbar d-flex align-items-center gap-1 px-2 py-1 border-bottom flex-wrap"
                    style={{ background: '#f3f4f6', flexShrink: 0 }}
                >
                    {toolbarActions.map((item, idx) =>
                        item.type === 'separator' ? (
                            <div key={idx} style={{ width: 1, height: 20, background: '#d1d5db', margin: '0 2px' }} />
                        ) : (
                            <button
                                key={idx}
                                className="btn btn-sm p-1"
                                style={{
                                    background: 'transparent',
                                    border: '1px solid transparent',
                                    borderRadius: '4px',
                                    color: '#374151',
                                    lineHeight: 1,
                                }}
                                title={item.title}
                                onClick={item.action}
                                onMouseEnter={(e) => { e.currentTarget.style.background = '#e5e7eb'; e.currentTarget.style.borderColor = '#9ca3af'; }}
                                onMouseLeave={(e) => { e.currentTarget.style.background = 'transparent'; e.currentTarget.style.borderColor = 'transparent'; }}
                            >
                                {item.icon}
                            </button>
                        )
                    )}
                </div>
            )}

            {/* Content area */}
            <div className="flex-grow-1 d-flex" style={{ overflow: 'hidden', minHeight: 0 }}>
                {/* Editor */}
                {(activeTab === 'edit' || activeTab === 'split') && (
                    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', borderRight: activeTab === 'split' ? '1px solid #e5e7eb' : 'none' }}>
                        <textarea
                            ref={textareaRef}
                            className="form-control border-0 rounded-0"
                            style={{
                                flex: 1,
                                resize: 'none',
                                fontFamily: "'JetBrains Mono', 'Fira Code', 'Cascadia Code', monospace",
                                fontSize: '13px',
                                lineHeight: '1.6',
                                padding: '12px 16px',
                                outline: 'none',
                                boxShadow: 'none',
                                background: '#fdfdfe',
                                color: '#1f2937',
                                overflow: 'auto',
                            }}
                            value={content}
                            onChange={(e) => setContent(e.target.value)}
                            onKeyDown={handleKeyDown}
                            onContextMenu={handleTextareaContextMenu}
                            placeholder="Write your markdown here...&#10;&#10;Equations: inline $E=mc^2$ or display $$\\int_0^\\infty e^{-x}\\,dx$$"
                            readOnly={readOnly}
                            spellCheck={false}
                        />
                    </div>
                )}

                {/* Preview */}
                {(activeTab === 'preview' || activeTab === 'split') && (
                    <div
                        ref={previewRef}
                        className="markdown-preview"
                        style={{
                            flex: 1,
                            overflow: 'auto',
                            padding: '12px 16px',
                            fontSize: '14px',
                            lineHeight: '1.7',
                            color: '#1f2937',
                            background: '#ffffff',
                        }}
                        dangerouslySetInnerHTML={getHtml()}
                    />
                )}
            </div>

            {/* Table insert popover */}
            {tablePopover && !readOnly && (
                <div
                    style={{
                        position: 'absolute',
                        top: 90,
                        right: 12,
                        background: '#ffffff',
                        border: '1px solid #d1d5db',
                        borderRadius: 8,
                        boxShadow: '0 4px 16px rgba(0,0,0,0.13)',
                        padding: '14px 16px',
                        zIndex: 2100,
                        minWidth: 200,
                    }}
                    onClick={e => e.stopPropagation()}
                >
                    <div className="fw-semibold mb-2" style={{ fontSize: 13 }}>Insert Table</div>
                    <div className="d-flex align-items-center gap-2 mb-2">
                        <label style={{ fontSize: 12, minWidth: 40 }}>Rows</label>
                        <input
                            type="number" min={2} max={20} value={tableRows}
                            onChange={e => setTableRows(Math.max(2, parseInt(e.target.value) || 2))}
                            className="form-control form-control-sm"
                            style={{ width: 70 }}
                        />
                    </div>
                    <div className="d-flex align-items-center gap-2 mb-3">
                        <label style={{ fontSize: 12, minWidth: 40 }}>Cols</label>
                        <input
                            type="number" min={1} max={10} value={tableCols}
                            onChange={e => setTableCols(Math.max(1, parseInt(e.target.value) || 1))}
                            className="form-control form-control-sm"
                            style={{ width: 70 }}
                        />
                    </div>
                    <div className="d-flex gap-2">
                        <button className="btn btn-sm btn-primary flex-grow-1" onClick={insertTable}
                            style={{ background: '#667eea', borderColor: '#667eea', fontSize: 12 }}>
                            Insert
                        </button>
                        <button className="btn btn-sm btn-outline-secondary" onClick={() => setTablePopover(false)}
                            style={{ fontSize: 12 }}>
                            Cancel
                        </button>
                    </div>
                </div>
            )}

            {/* Table context menu */}
            {tableMenu && (
                <>
                    <div
                        style={{ position: 'fixed', inset: 0, zIndex: 2199 }}
                        onClick={() => setTableMenu(null)}
                        onContextMenu={e => { e.preventDefault(); setTableMenu(null); }}
                    />
                    <div
                        style={{
                            position: 'fixed',
                            left: tableMenu.x,
                            top: tableMenu.y,
                            background: '#ffffff',
                            border: '1px solid #d1d5db',
                            borderRadius: 8,
                            boxShadow: '0 4px 16px rgba(0,0,0,0.15)',
                            zIndex: 2200,
                            minWidth: 190,
                            padding: '4px 0',
                            fontSize: 13,
                        }}
                        onClick={e => e.stopPropagation()}
                    >
                        {[
                            { label: '⬆ Insert Row Above', fn: () => tableOps.insertRowAbove(tableMenu.cursorPos) },
                            { label: '⬇ Insert Row Below', fn: () => tableOps.insertRowBelow(tableMenu.cursorPos) },
                            { label: '✕ Remove Row',        fn: () => tableOps.removeRow(tableMenu.cursorPos), danger: true },
                            { type: 'sep' },
                            { label: '⬅ Insert Column Left',  fn: () => tableOps.insertColLeft(tableMenu.cursorPos) },
                            { label: '➡ Insert Column Right', fn: () => tableOps.insertColRight(tableMenu.cursorPos) },
                            { label: '✕ Remove Column',       fn: () => tableOps.removeCol(tableMenu.cursorPos), danger: true },
                            { type: 'sep' },
                            { label: '⇔ Align Column Width',  fn: () => tableOps.prettifyCol(tableMenu.cursorPos), accent: true },
                            { label: '⊞ Prettify Whole Table', fn: () => tableOps.prettifyTable(tableMenu.cursorPos), accent: true },
                        ].map((item, idx) =>
                            item.type === 'sep' ? (
                                <div key={idx} style={{ height: 1, background: '#e5e7eb', margin: '3px 0' }} />
                            ) : (
                                <button
                                    key={idx}
                                    className="btn btn-sm w-100 text-start border-0 rounded-0 px-3 py-1"
                                    style={{
                                        color: item.danger ? '#dc2626' : item.accent ? '#7c3aed' : '#1f2937',
                                        background: 'transparent',
                                    }}
                                    onMouseEnter={e => e.currentTarget.style.background = item.danger ? '#fef2f2' : item.accent ? '#f5f3ff' : '#f3f4f6'}
                                    onMouseLeave={e => e.currentTarget.style.background = 'transparent'}
                                    onClick={() => { item.fn(); setTableMenu(null); }}
                                >
                                    {item.label}
                                </button>
                            )
                        )}
                    </div>
                </>
            )}

            {/* Asset Browser */}
            {assetBrowserOpen && (
                <AssetBrowser
                    project={project || ''}
                    onInsert={(snippet) => {
                        insertMarkdown(snippet);
                    }}
                    onClose={() => setAssetBrowserOpen(false)}
                />
            )}

            {/* Python File Browser */}
            {pythonFileBrowserOpen && (
                <PythonFileBrowser
                    onInsert={(snippet) => {
                        insertMarkdown(snippet);
                    }}
                    onClose={() => setPythonFileBrowserOpen(false)}
                />
            )}

            {/* Footer */}
            {/* discards changes on pressing cancel */}
            <div
                className="d-flex align-items-center justify-content-between px-3 py-2 border-top"
                style={{ background: '#f9fafb', flexShrink: 0 }}
            >
                <span className="text-muted" style={{ fontSize: '11px' }}>
                    {content.length} chars · {content.split(/\n/).length} lines
                </span>
                <div className="d-flex gap-2">
                    <button
                        className="btn btn-sm btn-outline-secondary"
                        onClick={() => { onClose(); }}
                    >
                        Cancel
                    </button>
                    {!readOnly && (
                        <button
                            className="btn btn-sm btn-primary"
                            onClick={() => onSave(content)}
                            style={{ background: '#667eea', borderColor: '#667eea' }}
                        >
                            Save
                        </button>
                    )}
                </div>
            </div>
        </div>
    );
}
