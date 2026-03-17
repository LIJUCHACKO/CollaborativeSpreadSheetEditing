import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, Upload, Image, Trash2, RefreshCw, Copy, Check, FolderOpen } from 'lucide-react';
import { authenticatedFetch, apiUrl } from '../utils/auth';

/**
 * AssetBrowser – modal that lets users browse / upload / delete images
 * stored in the project's `assets/` folder.
 *
 * Props:
 *   project   {string}   – current project path (e.g. "MyProject" or "MyProject/sub")
 *   onInsert  {function} – called with the markdown image snippet `![name](url)` to insert
 *   onClose   {function} – close the browser
 */
export default function AssetBrowser({ project, onInsert, onClose }) {
    const [assets, setAssets] = useState([]);
    const [loading, setLoading] = useState(false);
    const [uploading, setUploading] = useState(false);
    const [error, setError] = useState('');
    const [copiedName, setCopiedName] = useState(null);
    const [deleteConfirm, setDeleteConfirm] = useState(null); // asset name pending delete
    const [dragOver, setDragOver] = useState(false);
    const fileInputRef = useRef(null);

    const fetchAssets = useCallback(async () => {
        if (!project) return;
        setLoading(true);
        setError('');
        try {
            const res = await authenticatedFetch(
                apiUrl(`/api/assets?project=${encodeURIComponent(project)}`)
            );
            if (!res.ok) throw new Error(await res.text());
            const data = await res.json();
            setAssets(data || []);
        } catch (e) {
            setError('Failed to load assets: ' + e.message);
        } finally {
            setLoading(false);
        }
    }, [project]);

    useEffect(() => {
        fetchAssets();
    }, [fetchAssets]);

    const uploadFiles = async (files) => {
        if (!files || files.length === 0) return;
        setUploading(true);
        setError('');
        const results = [];
        for (const file of files) {
            const form = new FormData();
            form.append('file', file);
            try {
                const res = await authenticatedFetch(
                    apiUrl(`/api/assets?project=${encodeURIComponent(project)}`),
                    { method: 'POST', body: form }
                );
                if (!res.ok) {
                    const msg = await res.text();
                    setError(`Upload failed for "${file.name}": ${msg}`);
                } else {
                    const data = await res.json();
                    results.push(data);
                }
            } catch (e) {
                setError('Upload error: ' + e.message);
            }
        }
        setUploading(false);
        if (results.length > 0) fetchAssets();
    };

    const handleFileInput = (e) => {
        uploadFiles(Array.from(e.target.files));
        e.target.value = '';
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setDragOver(false);
        const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
        if (files.length === 0) {
            setError('Only image files can be uploaded.');
            return;
        }
        uploadFiles(files);
    };

    const handleDelete = async (name) => {
        try {
            const res = await authenticatedFetch(
                apiUrl(`/api/assets?project=${encodeURIComponent(project)}&name=${encodeURIComponent(name)}`),
                { method: 'DELETE' }
            );
            if (!res.ok) throw new Error(await res.text());
            setDeleteConfirm(null);
            fetchAssets();
        } catch (e) {
            setError('Delete failed: ' + e.message);
        }
    };

    const insertImage = (asset) => {
        const url = apiUrl(
            `/api/assets/serve?project=${encodeURIComponent(project)}&name=${encodeURIComponent(asset.name)}`
        );
        const label = asset.name.replace(/^\d+_/, ''); // strip timestamp prefix
        onInsert(`![${label}](${url})`);
        onClose();
    };

    const copyUrl = (asset) => {
        const url = apiUrl(
            `/api/assets/serve?project=${encodeURIComponent(project)}&name=${encodeURIComponent(asset.name)}`
        );
        navigator.clipboard.writeText(url).then(() => {
            setCopiedName(asset.name);
            setTimeout(() => setCopiedName(null), 1500);
        });
    };

    const formatSize = (bytes) => {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    };

    const displayName = (name) => name.replace(/^\d+_/, ''); // strip leading timestamp

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
                        width: 680,
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
                            background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                            color: '#fff',
                            padding: '12px 16px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            flexShrink: 0,
                        }}
                    >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <FolderOpen size={16} />
                            <span style={{ fontWeight: 600, fontSize: 14 }}>
                                Asset Browser — {project || '(no project)'}
                            </span>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                            <button
                                onClick={fetchAssets}
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
                            border: `2px dashed ${dragOver ? '#667eea' : '#d1d5db'}`,
                            borderRadius: 8,
                            background: dragOver ? '#f0f0ff' : '#fafafa',
                            padding: '14px 16px',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: 10,
                            transition: 'all 0.15s',
                            flexShrink: 0,
                        }}
                    >
                        <Upload size={18} color={dragOver ? '#667eea' : '#9ca3af'} />
                        <span style={{ fontSize: 13, color: dragOver ? '#667eea' : '#6b7280' }}>
                            {uploading
                                ? 'Uploading…'
                                : 'Click or drag & drop images here to upload (jpg, png, gif, webp, svg, bmp)'}
                        </span>
                        <input
                            ref={fileInputRef}
                            type="file"
                            accept="image/*"
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

                    {/* Asset grid */}
                    <div style={{
                        flex: 1,
                        overflowY: 'auto',
                        padding: '12px 16px 16px',
                    }}>
                        {loading && (
                            <div style={{ textAlign: 'center', color: '#9ca3af', padding: '32px 0', fontSize: 13 }}>
                                Loading assets…
                            </div>
                        )}
                        {!loading && assets.length === 0 && (
                            <div style={{ textAlign: 'center', padding: '40px 0' }}>
                                <Image size={40} color="#d1d5db" style={{ margin: '0 auto 12px', display: 'block' }} />
                                <div style={{ color: '#9ca3af', fontSize: 13 }}>No assets yet. Upload an image above.</div>
                            </div>
                        )}
                        {!loading && assets.length > 0 && (
                            <div style={{
                                display: 'grid',
                                gridTemplateColumns: 'repeat(auto-fill, minmax(150px, 1fr))',
                                gap: 12,
                            }}>
                                {assets.map(asset => (
                                    <div
                                        key={asset.name}
                                        style={{
                                            border: '1px solid #e5e7eb',
                                            borderRadius: 8,
                                            overflow: 'hidden',
                                            background: '#fff',
                                            display: 'flex',
                                            flexDirection: 'column',
                                            transition: 'box-shadow 0.15s',
                                            cursor: 'pointer',
                                        }}
                                        onMouseEnter={e => e.currentTarget.style.boxShadow = '0 2px 12px rgba(102,126,234,0.2)'}
                                        onMouseLeave={e => e.currentTarget.style.boxShadow = 'none'}
                                    >
                                        {/* Thumbnail */}
                                        <div
                                            style={{
                                                height: 110,
                                                background: '#f3f4f6',
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                overflow: 'hidden',
                                            }}
                                            onClick={() => insertImage(asset)}
                                            title="Click to insert"
                                        >
                                            <AssetThumbnail asset={asset} project={project} />
                                        </div>

                                        {/* Info + actions */}
                                        <div style={{ padding: '6px 8px', flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                                            <div
                                                style={{
                                                    fontSize: 11,
                                                    fontWeight: 600,
                                                    color: '#374151',
                                                    overflow: 'hidden',
                                                    textOverflow: 'ellipsis',
                                                    whiteSpace: 'nowrap',
                                                }}
                                                title={displayName(asset.name)}
                                            >
                                                {displayName(asset.name)}
                                            </div>
                                            <div style={{ fontSize: 10, color: '#9ca3af' }}>{formatSize(asset.size)}</div>
                                            <div style={{ display: 'flex', gap: 4, marginTop: 2 }}>
                                                {/* Insert button */}
                                                <button
                                                    onClick={() => insertImage(asset)}
                                                    title="Insert into markdown"
                                                    style={{
                                                        flex: 1,
                                                        fontSize: 10,
                                                        background: 'linear-gradient(135deg,#667eea,#764ba2)',
                                                        color: '#fff',
                                                        border: 'none',
                                                        borderRadius: 4,
                                                        padding: '3px 0',
                                                        cursor: 'pointer',
                                                    }}
                                                >
                                                    Insert
                                                </button>
                                                {/* Copy URL */}
                                                <button
                                                    onClick={() => copyUrl(asset)}
                                                    title="Copy URL"
                                                    style={{
                                                        background: '#f3f4f6',
                                                        border: '1px solid #e5e7eb',
                                                        borderRadius: 4,
                                                        padding: '3px 5px',
                                                        cursor: 'pointer',
                                                        display: 'flex',
                                                        alignItems: 'center',
                                                    }}
                                                >
                                                    {copiedName === asset.name
                                                        ? <Check size={11} color="#16a34a" />
                                                        : <Copy size={11} color="#6b7280" />}
                                                </button>
                                                {/* Delete */}
                                                {deleteConfirm === asset.name ? (
                                                    <>
                                                        <button
                                                            onClick={() => handleDelete(asset.name)}
                                                            title="Confirm delete"
                                                            style={{
                                                                background: '#fee2e2',
                                                                border: '1px solid #fca5a5',
                                                                borderRadius: 4,
                                                                padding: '3px 5px',
                                                                cursor: 'pointer',
                                                                fontSize: 10,
                                                                color: '#dc2626',
                                                            }}
                                                        >✓</button>
                                                        <button
                                                            onClick={() => setDeleteConfirm(null)}
                                                            style={{
                                                                background: '#f3f4f6',
                                                                border: '1px solid #e5e7eb',
                                                                borderRadius: 4,
                                                                padding: '3px 5px',
                                                                cursor: 'pointer',
                                                                fontSize: 10,
                                                                color: '#6b7280',
                                                            }}
                                                        >✗</button>
                                                    </>
                                                ) : (
                                                    <button
                                                        onClick={() => setDeleteConfirm(asset.name)}
                                                        title="Delete asset"
                                                        style={{
                                                            background: '#f3f4f6',
                                                            border: '1px solid #e5e7eb',
                                                            borderRadius: 4,
                                                            padding: '3px 5px',
                                                            cursor: 'pointer',
                                                            display: 'flex',
                                                            alignItems: 'center',
                                                        }}
                                                    >
                                                        <Trash2 size={11} color="#6b7280" />
                                                    </button>
                                                )}
                                            </div>
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

/** Lazy thumbnail with auth – fetches the image using the auth token and renders via object URL */
function AssetThumbnail({ asset, project }) {
    const [src, setSrc] = useState(null);
    useEffect(() => {
        let revoked = false;
        fetch(apiUrl(`/api/assets/serve?project=${encodeURIComponent(project)}&name=${encodeURIComponent(asset.name)}`))
            .then(r => r.blob())
            .then(blob => {
                if (!revoked) setSrc(URL.createObjectURL(blob));
            })
            .catch(() => {});
        return () => {
            revoked = true;
            if (src) URL.revokeObjectURL(src);
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [asset.name, project]);

    if (!src) {
        return <Image size={32} color="#d1d5db" />;
    }
    return (
        <img
            src={src}
            alt={asset.name}
            style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
        />
    );
}
