import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, Upload, Trash2, RefreshCw, Copy, Check, FolderOpen, FileCode, Link, Image } from 'lucide-react';
import { authenticatedFetch, apiUrl } from '../utils/auth';

/**
 * PythonFileBrowser – modal that lets users browse / upload / delete files
 * stored in the shared `pythonDirectory`.
 *
 * Props:
 *   onInsert  {function} – called with a markdown link snippet `[name](url)` to insert
 *   onClose   {function} – close the browser
 */
export default function PythonFileBrowser({ onInsert, onClose }) {
    const [files, setFiles] = useState([]);
    const [loading, setLoading] = useState(false);
    const [uploading, setUploading] = useState(false);
    const [error, setError] = useState('');
    const [copiedName, setCopiedName] = useState(null);
    const [deleteConfirm, setDeleteConfirm] = useState(null);
    const [dragOver, setDragOver] = useState(false);
    const fileInputRef = useRef(null);

    const fetchFiles = useCallback(async () => {
        setLoading(true);
        setError('');
        try {
            const res = await authenticatedFetch(apiUrl('/api/python-files'));
            if (!res.ok) throw new Error(await res.text());
            const data = await res.json();
            setFiles(data || []);
        } catch (e) {
            setError('Failed to load files: ' + e.message);
        } finally {
            setLoading(false);
        }
    }, []);

    useEffect(() => {
        fetchFiles();
    }, [fetchFiles]);

    const uploadFiles = async (selectedFiles) => {
        if (!selectedFiles || selectedFiles.length === 0) return;
        setUploading(true);
        setError('');
        const results = [];
        for (const file of selectedFiles) {
            const form = new FormData();
            form.append('file', file);
            try {
                const res = await authenticatedFetch(
                    apiUrl('/api/python-files'),
                    { method: 'POST', body: form }
                );
                if (!res.ok) {
                    const msg = await res.text();
                    setError(`Upload failed for "${file.name}": ${msg}`);
                } else {
                    results.push(await res.json());
                }
            } catch (e) {
                setError('Upload error: ' + e.message);
            }
        }
        setUploading(false);
        if (results.length > 0) fetchFiles();
    };

    const handleFileInput = (e) => {
        uploadFiles(Array.from(e.target.files));
        e.target.value = '';
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setDragOver(false);
        uploadFiles(Array.from(e.dataTransfer.files));
    };

    const handleDelete = async (name) => {
        try {
            const res = await authenticatedFetch(
                apiUrl(`/api/python-files?name=${encodeURIComponent(name)}`),
                { method: 'DELETE' }
            );
            if (!res.ok) throw new Error(await res.text());
            setDeleteConfirm(null);
            fetchFiles();
        } catch (e) {
            setError('Delete failed: ' + e.message);
        }
    };

    const IMAGE_EXTS = new Set(['jpg', 'jpeg', 'png', 'gif', 'webp', 'svg', 'bmp']);
    const isImage = (name) => IMAGE_EXTS.has(name.split('.').pop().toLowerCase());

    const insertLink = (file) => {
        const url = apiUrl(`/api/python-files/serve?name=${encodeURIComponent(file.name)}`);
        // Use image syntax for image files, plain link for everything else
        onInsert(isImage(file.name) ? `![${file.name}](${url})` : `[${file.name}](${url})`);
        onClose();
    };

    const copyUrl = (file) => {
        const url = apiUrl(`/api/python-files/serve?name=${encodeURIComponent(file.name)}`);
        navigator.clipboard.writeText(url).then(() => {
            setCopiedName(file.name);
            setTimeout(() => setCopiedName(null), 1500);
        });
    };

    const formatSize = (bytes) => {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    };

    /** Pick a colour accent based on file extension */
    const extColor = (name) => {
        const ext = name.split('.').pop().toLowerCase();
        const map = {
            py: '#3b82f6', js: '#f59e0b', ts: '#2563eb',
            json: '#10b981', csv: '#8b5cf6', txt: '#6b7280',
            sh: '#f97316', md: '#ec4899',
        };
        return map[ext] || '#9ca3af';
    };

    return (
        <>
            {/* Backdrop */}
            <div
                style={{
                    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
                    zIndex: 3000, display: 'flex', alignItems: 'center', justifyContent: 'center',
                }}
                onClick={onClose}
            >
                {/* Modal */}
                <div
                    style={{
                        background: '#fff',
                        borderRadius: 12,
                        boxShadow: '0 8px 40px rgba(0,0,0,0.25)',
                        width: 640,
                        maxWidth: '95vw',
                        maxHeight: '85vh',
                        display: 'flex',
                        flexDirection: 'column',
                        overflow: 'hidden',
                    }}
                    onClick={e => e.stopPropagation()}
                >
                    {/* Header */}
                    <div
                        style={{
                            background: 'linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%)',
                            color: '#fff',
                            padding: '12px 16px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            flexShrink: 0,
                        }}
                    >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <FileCode size={16} />
                            <span style={{ fontWeight: 600, fontSize: 14 }}>
                                Python File Browser — pythonDirectory
                            </span>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                            <button
                                onClick={fetchFiles}
                                title="Refresh"
                                style={{ background: 'transparent', border: 'none', color: '#fff', cursor: 'pointer', padding: 4, display: 'flex' }}
                            >
                                <RefreshCw size={14} />
                            </button>
                            <button
                                onClick={onClose}
                                title="Close"
                                style={{ background: 'transparent', border: 'none', color: '#fff', cursor: 'pointer', padding: 4, display: 'flex' }}
                            >
                                <X size={16} />
                            </button>
                        </div>
                    </div>

                    {/* Upload zone */}
                    <div
                        onDrop={handleDrop}
                        onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                        onDragLeave={() => setDragOver(false)}
                        onClick={() => fileInputRef.current?.click()}
                        style={{
                            margin: '12px 16px 0',
                            border: `2px dashed ${dragOver ? '#3b82f6' : '#d1d5db'}`,
                            borderRadius: 8,
                            background: dragOver ? '#eff6ff' : '#fafafa',
                            padding: '14px 16px',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: 10,
                            transition: 'all 0.15s',
                            flexShrink: 0,
                        }}
                    >
                        <Upload size={18} color={dragOver ? '#3b82f6' : '#9ca3af'} />
                        <span style={{ fontSize: 13, color: dragOver ? '#3b82f6' : '#6b7280' }}>
                            {uploading ? 'Uploading…' : 'Click or drag & drop files here to upload'}
                        </span>
                        <input
                            ref={fileInputRef}
                            type="file"
                            multiple
                            style={{ display: 'none' }}
                            onChange={handleFileInput}
                        />
                    </div>

                    {/* Error */}
                    {error && (
                        <div style={{
                            margin: '8px 16px 0',
                            background: '#fef2f2',
                            border: '1px solid #fca5a5',
                            borderRadius: 6,
                            padding: '8px 12px',
                            fontSize: 12,
                            color: '#dc2626',
                            flexShrink: 0,
                        }}>
                            {error}
                            <button
                                onClick={() => setError('')}
                                style={{ float: 'right', background: 'none', border: 'none', cursor: 'pointer', color: '#dc2626', fontSize: 14 }}
                            >×</button>
                        </div>
                    )}

                    {/* File list */}
                    <div style={{ flex: 1, overflowY: 'auto', padding: '12px 16px 16px' }}>
                        {loading && (
                            <div style={{ textAlign: 'center', color: '#9ca3af', padding: '32px 0', fontSize: 13 }}>
                                Loading files…
                            </div>
                        )}
                        {!loading && files.length === 0 && (
                            <div style={{ textAlign: 'center', padding: '40px 0' }}>
                                <FileCode size={40} color="#d1d5db" style={{ margin: '0 auto 12px', display: 'block' }} />
                                <div style={{ color: '#9ca3af', fontSize: 13 }}>No files yet. Upload a file above.</div>
                            </div>
                        )}
                        {!loading && files.length > 0 && (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                {files.map(file => (
                                    <div
                                        key={file.name}
                                        style={{
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: 10,
                                            padding: '8px 12px',
                                            border: '1px solid #e5e7eb',
                                            borderRadius: 8,
                                            background: '#fff',
                                            transition: 'box-shadow 0.15s',
                                        }}
                                        onMouseEnter={e => e.currentTarget.style.boxShadow = '0 2px 8px rgba(59,130,246,0.15)'}
                                        onMouseLeave={e => e.currentTarget.style.boxShadow = 'none'}
                                    >
                                        {/* Thumbnail for images, extension badge for other files */}
                                        <div style={{
                                            width: 48, height: 48, borderRadius: 6, flexShrink: 0,
                                            background: extColor(file.name) + '22',
                                            display: 'flex', alignItems: 'center', justifyContent: 'center',
                                            overflow: 'hidden',
                                        }}>
                                            {isImage(file.name)
                                                ? <PythonFileThumbnail name={file.name} />
                                                : <span style={{
                                                    fontSize: 9, fontWeight: 700, color: extColor(file.name),
                                                    textTransform: 'uppercase', letterSpacing: 0.5,
                                                }}>
                                                    {file.name.split('.').pop().slice(0, 4)}
                                                </span>
                                            }
                                        </div>

                                        {/* Name + size */}
                                        <div style={{ flex: 1, minWidth: 0 }}>
                                            <div style={{
                                                fontSize: 13, fontWeight: 600, color: '#1f2937',
                                                overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                                            }} title={file.name}>
                                                {file.name}
                                            </div>
                                            <div style={{ fontSize: 11, color: '#9ca3af' }}>{formatSize(file.size)}</div>
                                        </div>

                                        {/* Actions */}
                                        <div style={{ display: 'flex', gap: 4, flexShrink: 0 }}>
                                            {/* Insert link */}
                                            <button
                                                onClick={() => insertLink(file)}
                                                title="Insert link into markdown"
                                                style={{
                                                    fontSize: 11,
                                                    background: 'linear-gradient(135deg,#3b82f6,#1d4ed8)',
                                                    color: '#fff',
                                                    border: 'none',
                                                    borderRadius: 4,
                                                    padding: '4px 10px',
                                                    cursor: 'pointer',
                                                    display: 'flex', alignItems: 'center', gap: 4,
                                                }}
                                            >
                                                <Link size={11} /> Insert
                                            </button>

                                            {/* Copy URL */}
                                            <button
                                                onClick={() => copyUrl(file)}
                                                title="Copy URL"
                                                style={{
                                                    background: '#f3f4f6',
                                                    border: '1px solid #e5e7eb',
                                                    borderRadius: 4,
                                                    padding: '4px 7px',
                                                    cursor: 'pointer',
                                                    display: 'flex', alignItems: 'center',
                                                }}
                                            >
                                                {copiedName === file.name
                                                    ? <Check size={12} color="#16a34a" />
                                                    : <Copy size={12} color="#6b7280" />}
                                            </button>

                                            {/* Delete */}
                                            {deleteConfirm === file.name ? (
                                                <>
                                                    <button
                                                        onClick={() => handleDelete(file.name)}
                                                        title="Confirm delete"
                                                        style={{
                                                            background: '#fee2e2', border: '1px solid #fca5a5',
                                                            borderRadius: 4, padding: '4px 7px',
                                                            cursor: 'pointer', fontSize: 11, color: '#dc2626',
                                                        }}
                                                    >✓</button>
                                                    <button
                                                        onClick={() => setDeleteConfirm(null)}
                                                        style={{
                                                            background: '#f3f4f6', border: '1px solid #e5e7eb',
                                                            borderRadius: 4, padding: '4px 7px',
                                                            cursor: 'pointer', fontSize: 11, color: '#6b7280',
                                                        }}
                                                    >✗</button>
                                                </>
                                            ) : (
                                                <button
                                                    onClick={() => setDeleteConfirm(file.name)}
                                                    title="Delete file"
                                                    style={{
                                                        background: '#f3f4f6', border: '1px solid #e5e7eb',
                                                        borderRadius: 4, padding: '4px 7px',
                                                        cursor: 'pointer', display: 'flex', alignItems: 'center',
                                                    }}
                                                >
                                                    <Trash2 size={12} color="#6b7280" />
                                                </button>
                                            )}
                                        </div>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </>
    );
}

/** Thumbnail for image files stored in pythonDirectory.
 *  The /api/python-files/serve endpoint is public (no auth needed),
 *  so we can use a plain <img> tag with the direct URL. */
function PythonFileThumbnail({ name }) {
    const src = apiUrl(`/api/python-files/serve?name=${encodeURIComponent(name)}`);
    const [ok, setOk] = React.useState(true);
    if (!ok) return <Image size={24} color="#d1d5db" />;
    return (
        <img
            src={src}
            alt={name}
            onError={() => setOk(false)}
            style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
        />
    );
}
