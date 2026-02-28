import React, { useState, useRef, useEffect, useCallback } from 'react';
import { marked } from 'marked';
import { X, Eye, Edit3, Bold, Italic, Heading1, Heading2, Heading3, List, ListOrdered, Code, Link, Image, Quote, Minus, CheckSquare, Maximize2, Minimize2, Sigma } from 'lucide-react';

// Configure marked for safe rendering
marked.setOptions({
    breaks: true,
    gfm: true,
});

// Load MathJax from CDN once per page lifetime
function ensureMathJax() {
    if (typeof window === 'undefined') return;
    if (window.MathJax) return; // already loaded or loading
    // Configure MathJax before the script loads
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
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js';
    script.async = true;
    script.id = 'mathjax-cdn-script';
    document.head.appendChild(script);
}

export default function MarkdownEditorPanel({ cellRow, cellCol, value, onSave, onClose, readOnly }) {
    const [content, setContent] = useState(value || '');
    const [activeTab, setActiveTab] = useState('edit'); // 'edit' | 'preview' | 'split'
    const [isMaximized, setIsMaximized] = useState(false);
    const textareaRef = useRef(null);
    const panelRef = useRef(null);
    const previewRef = useRef(null);
    const [isDragging, setIsDragging] = useState(false);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [position, setPosition] = useState({ x: null, y: null });

    // Ensure MathJax CDN is loaded
    useEffect(() => { ensureMathJax(); }, []);

    // Sync incoming value when cell changes
    useEffect(() => {
        setContent(value || '');
    }, [value, cellRow, cellCol]);

    // Position panel on mount
    useEffect(() => {
        if (position.x === null) {
            const vw = window.innerWidth;
            const panelWidth = isMaximized ? vw * 0.8 : 520;
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
    ];

    // Keyboard shortcuts
    const handleKeyDown = (e) => {
        if (e.ctrlKey || e.metaKey) {
            if (e.key === 'b') { e.preventDefault(); insertMarkdown('**', '**', 'bold text'); }
            if (e.key === 'i') { e.preventDefault(); insertMarkdown('*', '*', 'italic text'); }
            if (e.key === 's') { e.preventDefault(); onSave(content); }
        }
    };

    const getHtml = () => {
        try {
            return { __html: marked.parse(content || '') };
        } catch {
            return { __html: '<p class="text-danger">Error rendering markdown</p>' };
        }
    };

    const panelWidth = isMaximized ? '80vw' : '520px';
    const panelHeight = isMaximized ? '80vh' : '480px';

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
                minHeight: '320px',
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
                        onClick={onClose}
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

            {/* Footer */}
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
                        onClick={onClose}
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
