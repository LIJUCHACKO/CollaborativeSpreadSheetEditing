import React, { useState, useRef, useEffect, useCallback } from 'react';
import { X, Maximize2, Minimize2, Sparkles, Play, CornerDownLeft, Eye } from 'lucide-react';
import { authenticatedFetch, apiUrl } from '../utils/auth';

/**
 * Highlight {{CellRef}} template references and plain text in an AI prompt.
 * Returns an array of { lineNum, tokens } objects.
 */
function highlightPrompt(code) {
    const lines = (code || '').split('\n');
    return lines.map((line, li) => {
        const tokens = [];
        let i = 0;
        while (i < line.length) {
            // Template references {{...}}
            if (line[i] === '{' && line[i + 1] === '{') {
                const end = line.indexOf('}}', i + 2);
                const refEnd = end !== -1 ? end + 2 : line.length;
                tokens.push(
                    <span key={`${li}-${i}`} style={{ color: '#4fc1ff', fontWeight: 'bold' }}>
                        {line.substring(i, refEnd)}
                    </span>
                );
                i = refEnd;
                continue;
            }
            // Collect plain text until the next {{ 
            let j = i + 1;
            while (j < line.length && !(line[j] === '{' && line[j + 1] === '{')) j++;
            tokens.push(
                <span key={`${li}-${i}`} style={{ color: '#d4d4d4' }}>
                    {line.substring(i, j)}
                </span>
            );
            i = j;
        }
        return { lineNum: li + 1, tokens };
    });
}

/**
 * AIPromptEditorPanel — a floating, draggable AI prompt editor panel
 * modeled after ScriptEditorPanel.
 *
 * Props:
 *  - cellRow, cellCol: cell coordinates
 *  - aiPromptText, setAIPromptText: controlled text state
 *  - canEdit: boolean
 *  - onApply: () => void — called when Apply is clicked
 *  - onClose: () => void — called when Cancel / X is clicked
 *  - onInsertRange: () => void — called when Insert Range is clicked
 *  - textareaRef: ref to pass to the internal textarea
 */
export default function AIPromptEditorPanel({
    cellRow,
    cellCol,
    projectName,
    sheetName,
    aiPromptText,
    setAIPromptText,
    canEdit,
    onApply,
    onClose,
    onInsertRange,
    textareaRef: externalTextareaRef,
}) {
    const panelRef = useRef(null);
    const internalTextareaRef = useRef(null);
    const textareaRef = externalTextareaRef || internalTextareaRef;

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
            apiUrl(`/api/preview/ai-prompt?project=${encodeURIComponent(projectName || '')}&sheet=${encodeURIComponent(sheetName || '')}`),
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ prompt: aiPromptText || '' }),
            }
        )
            .then(r => r.ok ? r.json() : r.text().then(t => Promise.reject(t)))
            .then(data => { setPreviewText(data.resolved ?? ''); })
            .catch(err => { setPreviewError(String(err)); })
            .finally(() => setPreviewLoading(false));
    }, [activeTab, aiPromptText, projectName, sheetName]);

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
        if (e.target.closest('.ai-prompt-header')) {
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

    const lineCount = (aiPromptText || '').split('\n').length;
    const highlighted = highlightPrompt(aiPromptText || '');

    const panelWidth = isMaximized ? '85vw' : '680px';
    const panelHeight = isMaximized ? '80vh' : '480px';

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
                minHeight: '300px',
            }}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
        >
            {/* Header — draggable */}
            <div
                className="ai-prompt-header d-flex align-items-center justify-content-between px-3 py-2"
                style={{
                    background: 'linear-gradient(135deg, #6a1fa2 0%, #4a1275 100%)',
                    color: 'white',
                    cursor: isDragging ? 'grabbing' : 'grab',
                    userSelect: 'none',
                    flexShrink: 0,
                }}
                onMouseDown={handleMouseDown}
            >
                <div className="d-flex align-items-center gap-2">
                    <Sparkles size={16} />
                    <span className="fw-semibold small">
                        AI Prompt — Cell {cellCol}{cellRow}
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
                    className="btn btn-sm px-3 py-1 rounded-0 border-0"
                    style={{
                        borderBottom: activeTab === 'edit' ? '2px solid #9b4dca' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'edit' ? '#c792ea' : '#888',
                        fontWeight: activeTab === 'edit' ? 'bold' : 'normal',
                    }}
                    onClick={() => setActiveTab('edit')}
                >
                    <Sparkles size={12} className="me-1" /> Edit
                </button>
                <button
                    className="btn btn-sm px-3 py-1 rounded-0 border-0"
                    style={{
                        borderBottom: activeTab === 'highlight' ? '2px solid #9b4dca' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'highlight' ? '#c792ea' : '#888',
                        fontWeight: activeTab === 'highlight' ? 'bold' : 'normal',
                    }}
                    onClick={() => setActiveTab('highlight')}
                >
                    <Sparkles size={12} className="me-1" /> Highlighted
                </button>
                <button
                    className="btn btn-sm px-3 py-1 rounded-0 border-0"
                    style={{
                        borderBottom: activeTab === 'preview' ? '2px solid #9b4dca' : '2px solid transparent',
                        background: 'transparent',
                        color: activeTab === 'preview' ? '#c792ea' : '#888',
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
                <div
                    className="d-flex align-items-center gap-3 px-3 py-1"
                    style={{ background: '#2d2d2d', borderBottom: '1px solid #333', flexShrink: 0 }}
                >
                    <label
                        className="d-flex align-items-center gap-1 small"
                        style={{ color: '#aaa', cursor: 'pointer', fontSize: '0.75rem' }}
                    >
                        <input
                            type="checkbox"
                            checked={showLineNumbers}
                            onChange={(e) => setShowLineNumbers(e.target.checked)}
                            style={{ accentColor: '#9b4dca' }}
                        />
                        Line #
                    </label>
                    <div
                        className="d-flex align-items-center gap-1 ms-auto"
                        style={{ fontSize: '0.7rem', color: '#666' }}
                    >
                        <span title="Use {{A1}} or {{A1:B3}} to reference cell ranges in your prompt.">
                            Use {'{{A1}}'} or {'{{A1:B3}}'} for cell references
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
                                        if (el && textareaRef.current) {
                                            const ta = textareaRef.current;
                                            const handler = () => { el.scrollTop = ta.scrollTop; };
                                            ta.removeEventListener('scroll', handler);
                                            ta.addEventListener('scroll', handler);
                                        }
                                    }}
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
                                    whiteSpace: 'pre-wrap',
                                    overflowWrap: 'break-word',
                                    overflowX: 'hidden',
                                    caretColor: '#aeafad',
                                }}
                                spellCheck={false}
                                value={aiPromptText}
                                onChange={(e) => setAIPromptText(e.target.value)}
                                disabled={!canEdit}
                                placeholder={`Enter AI prompt for cell ${cellCol}${cellRow}.\nUse {{A1}} or {{A1:B3}} to reference cell values.`}
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
                                    <span style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
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
                                overflowX: 'hidden',
                                background: '#1e1e1e',
                                padding: '12px 16px',
                                color: '#d4d4d4',
                                fontFamily: 'system-ui, sans-serif',
                                fontSize: '13px',
                                lineHeight: '1.6',
                                whiteSpace: 'pre-wrap',
                                wordBreak: 'break-word',
                            }}
                        >
                            {previewLoading && (
                                <span style={{ color: '#888', fontStyle: 'italic' }}>Resolving references…</span>
                            )}
                            {!previewLoading && previewError && (
                                <span style={{ color: '#f48771' }}>Error: {previewError}</span>
                            )}
                            {!previewLoading && !previewError && (
                                previewText
                                    ? previewText
                                    : <span style={{ color: '#555', fontStyle: 'italic' }}>Nothing to preview.</span>
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
                            background: '#6a1fa2',
                            color: 'white',
                            border: 'none',
                            fontSize: '0.8rem',
                        }}
                        onClick={onInsertRange}
                        disabled={!canEdit}
                        title="Insert selected range into prompt at cursor position"
                    >
                        <CornerDownLeft size={13} className="me-1" />
                        Insert Range
                    </button>
                    <span style={{ color: '#666', fontSize: '0.7rem' }}>
                        {lineCount} line{lineCount !== 1 ? 's' : ''} · {(aiPromptText || '').length} chars
                    </span>
                </div>
                <div className="d-flex align-items-center gap-2">
                    <button
                        className="btn btn-sm"
                        style={{
                            background: '#6a1fa2',
                            color: 'white',
                            border: 'none',
                            fontSize: '0.8rem',
                        }}
                        onClick={onApply}
                        disabled={!canEdit}
                        title="Apply AI prompt"
                    >
                        <Play size={13} className="me-1" />
                        Apply
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
