import React, { useEffect, useState, useRef } from 'react';
import { useNavigate, useParams, Link } from 'react-router-dom';
import {
    FileSpreadsheet,
    Search,
    LogOut,
    User,
    MoreVertical,
    Trash2,
    Edit2,
    History,
    ArrowLeft,
    Copy,
    ClipboardPaste,
    X
} from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';
import { isSessionValid, clearAuth, authenticatedFetch, getUsername } from '../utils/auth';

// Shared clipboard helpers using localStorage
function getClipboard() {
    try {
        const raw = localStorage.getItem('app_clipboard');
        return raw ? JSON.parse(raw) : null;
    } catch { return null; }
}
function setClipboardStorage(data) {
    if (data) localStorage.setItem('app_clipboard', JSON.stringify(data));
    else localStorage.removeItem('app_clipboard');
}

export default function Dashboard() {
    const { project } = useParams();
    const navigate = useNavigate();
    const username = getUsername();

    // Sheets and UI state
    const [sheets, setSheets] = useState([]);
    const [newSheetName, setNewSheetName] = useState('');
    const [newSheetType, setNewSheetType] = useState('datasheet'); // 'datasheet' or 'document'
    const [isCreating, setIsCreating] = useState(true);
    const [searchQuery, setSearchQuery] = useState('');
    const [editingSheetId, setEditingSheetId] = useState(null);
    const [editingSheetName, setEditingSheetName] = useState('');

    // Shared clipboard: {type:'folder'|'sheet', sourcePath, sourceSheetId?}
    const [clipboard, setClipboard] = useState(() => getClipboard());
    const [pastingTarget, setPastingTarget] = useState(null); // '__here__' for current folder, or subfolder name
    const [pasteName, setPasteName] = useState('');

    // Project owner (only top-level project owner can create/paste sheets & subfolders)
    const [projectOwner, setProjectOwner] = useState('');

    // Audit sidebar state
    const [auditLog, setAuditLog] = useState([]);
    const [isAuditOpen, setAuditOpen] = useState(false);
    const [selectedAuditId, setSelectedAuditId] = useState(null);
    const auditLogRef = useRef(null);
    const auditLogScrollTopRef = useRef(0);
    const fileInputRef = useRef(null);

    // Folder navigation state
    const [currentPath, setCurrentPath] = useState(project || '');
    const [folders, setFolders] = useState([]);
    const [newFolderName, setNewFolderName] = useState('');
    const [editingFolderName, setEditingFolderName] = useState(null);
    const [editingFolderNewName, setEditingFolderNewName] = useState('');

    // Effects
    useEffect(() => {
        if (!username || !isSessionValid()) {
            clearAuth();
            navigate('/');
            return;
        }
        setCurrentPath(project || '');
        fetchSheets(project || '');
        if (project) fetchFolders(project);
        if (project) fetchProjectOwner(project);

        const sessionCheckInterval = setInterval(() => {
            if (!isSessionValid()) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        }, 60000);
        return () => clearInterval(sessionCheckInterval);
    }, [project, username, navigate]);

    // Sync clipboard from localStorage (cross-tab and same-tab navigation)
    useEffect(() => {
        const syncClipboard = () => setClipboard(getClipboard());
        window.addEventListener('storage', syncClipboard);
        window.addEventListener('focus', syncClipboard);
        return () => {
            window.removeEventListener('storage', syncClipboard);
            window.removeEventListener('focus', syncClipboard);
        };
    }, []);

    useEffect(() => {
        if (isAuditOpen && auditLogRef.current) {
            auditLogRef.current.scrollTop = auditLogScrollTopRef.current;
        }
    }, [isAuditOpen]);

    // Data fetchers
    const fetchSheets = async (pathOverride) => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const path = typeof pathOverride === 'string' ? pathOverride : currentPath;
            const query = path ? `?project=${encodeURIComponent(path)}` : '';
            const res = await authenticatedFetch(`http://${host}:8082/api/sheets${query}`);
            if (res.ok) {
                const data = await res.json();
                setSheets(data || []);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        } catch (error) {
            console.error('Failed to fetch sheets', error);
        }
    };

    const fetchFolders = async (pathOverride) => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const path = typeof pathOverride === 'string' ? pathOverride : currentPath;
            if (!path) { setFolders([]); return; }
            const res = await authenticatedFetch(`http://${host}:8082/api/folders?project=${encodeURIComponent(path)}`);
            if (res.ok) {
                const data = await res.json();
                const names = Array.isArray(data) ? data.map(f => f.name) : [];
                setFolders(names);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                setFolders([]);
            }
        } catch (error) {
            console.error('Failed to fetch folders', error);
            setFolders([]);
        }
    };

    const fetchProjectOwner = async (pathOverride) => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const topProject = ((typeof pathOverride === 'string' ? pathOverride : project) || '').split('/')[0];
            if (!topProject) { setProjectOwner(''); return; }
            const res = await authenticatedFetch(`http://${host}:8082/api/projects`);
            if (res.ok) {
                const list = await res.json();
                const found = Array.isArray(list) ? list.find(p => p.name === topProject) : null;
                setProjectOwner(found?.owner || '');
            }
        } catch (e) { /* ignore */ }
    };

    const fetchProjectAudit = async () => {
        try {
            const topProject = (project || '').split('/')[0];
            if (!topProject) { setAuditLog([]); return; }
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8082/api/projects/audit?project=${encodeURIComponent(topProject)}`);
            if (res.ok) {
                const entries = await res.json();
                setAuditLog(Array.isArray(entries) ? entries : []);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        } catch (e) {
            // ignore fetch errors
        }
    };

    // Actions
    const createSheet = async (e) => {
        e.preventDefault();
        if (!newSheetName.trim()) return;
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const body = currentPath
                ? { name: newSheetName, user: username, project_name: currentPath, sheet_type: newSheetType }
                : { name: newSheetName, user: username, sheet_type: newSheetType };
            const res = await authenticatedFetch(`http://${host}:8082/api/sheets`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.ok) {
                const sheet = await res.json();
                setNewSheetName('');
                setNewSheetType('datasheet');
                setIsCreating(false);
                fetchSheets();
                const path = currentPath;
                const route = newSheetType === 'document' ? 'document' : 'sheet';
                navigate(path ? `/${route}/${sheet.name}?project=${encodeURIComponent(path)}` : `/${route}/${sheet.name}`);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to create sheet');
            }
        } catch (error) {
            console.error('Failed to create sheet', error);
        }
    };

    const handleLogout = async () => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            await authenticatedFetch(`http://${host}:8082/api/logout`, { method: 'POST' });
        } catch (error) {
            console.error('Logout error', error);
        } finally {
            clearAuth();
            navigate('/');
        }
    };

    const handleDownloadProjectXlsx = async () => {
        try {
            const path = currentPath || project;
            if (!path) return;
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8082/api/export_project?project=${encodeURIComponent(path)}`, { method: 'GET' });
            if (!res.ok) {
                const text = await res.text();
                alert(`Failed to export project: ${text}`);
                return;
            }
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${path}.xlsx`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Error downloading project XLSX', err);
            alert('An unexpected error occurred while exporting the project.');
        }
    };

    const handleImportProjectXlsx = async (file) => {
        try {
            const path = currentPath || project;
            if (!path || !file) return;
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const form = new FormData();
            form.append('file', file);
            const res = await authenticatedFetch(`http://${host}:8082/api/import_project_xlsx?project=${encodeURIComponent(path)}`, {
                method: 'POST',
                body: form,
            });
            if (!res.ok) {
                const text = await res.text();
                alert(text || 'Failed to import XLSX');
                return;
            }
            const data = await res.json();
            const count = Array.isArray(data?.created) ? data.created.length : 0;
            alert(`Imported ${count} sheet(s) from XLSX`);
            fetchSheets();
        } catch (err) {
            console.error('Error importing project XLSX', err);
            alert('An unexpected error occurred while importing the XLSX.');
        } finally {
            if (fileInputRef.current) fileInputRef.current.value = '';
        }
    };

    // Folder helpers
    const goToFolder = (name) => {
        const base = currentPath || '';
        const next = base ? `${base}/${name}` : name;
        setCurrentPath(next);
        navigate(`/project/${encodeURIComponent(next)}`);
        fetchSheets(next);
        fetchFolders(next);
        fetchProjectOwner(next);
    };

    const goUpOne = () => {
        const parts = (currentPath || '').split('/').filter(Boolean);
        if (parts.length <= 1) {
            const top = parts[0] || '';
            navigate('/projects');
            setCurrentPath(top);
            fetchSheets(top);
            fetchFolders(top);
            fetchProjectOwner(top);
            return;
        }
        const next = parts.slice(0, -1).join('/');
        setCurrentPath(next);
        navigate(`/project/${encodeURIComponent(next)}`);
        fetchSheets(next);
        fetchFolders(next);
        fetchProjectOwner(next);
    };

    const createFolder = async (e) => {
        e.preventDefault();
        const name = newFolderName.trim();
        if (!name) return;
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const body = { parent: currentPath || project || '', name };
            const res = await authenticatedFetch(`http://${host}:8082/api/folders`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.ok) {
                setNewFolderName('');
                fetchFolders();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to create folder');
            }
        } catch (error) {
            console.error('Failed to create folder', error);
        }
    };

    const startRenamingFolder = (name) => {
        setEditingFolderName(name);
        setEditingFolderNewName(name);
    };

    const cancelRenamingFolder = () => {
        setEditingFolderName(null);
        setEditingFolderNewName('');
    };

    const renameFolder = async (oldName) => {
        const newName = editingFolderNewName.trim();
        if (!newName || newName === oldName) {
            cancelRenamingFolder();
            return;
        }
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const body = { parent: currentPath || project || '', old_name: oldName, new_name: newName };
            const res = await authenticatedFetch(`http://${host}:8082/api/folders`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.status === 409) {
                const errorText = await res.text();
                alert(errorText || 'Cannot rename: one or more sheets in this folder are currently open by users.');
                return;
            }
            if (res.ok) {
                cancelRenamingFolder();
                fetchFolders();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else if (res.status === 403) {
                alert('Only the project owner can rename folders.');
            } else {
                const text = await res.text();
                alert(text || 'Failed to rename folder');
            }
        } catch (error) {
            console.error('Failed to rename folder', error);
            alert('Error renaming folder');
        }
    };

    // --- Unified clipboard (shared via localStorage) ---
    const handleCopyFolder = (folderName) => {
        const base = currentPath || '';
        const fullPath = base ? `${base}/${folderName}` : folderName;
        const data = { type: 'folder', sourcePath: fullPath };
        setClipboard(data);
        setClipboardStorage(data);
        cancelPaste();
    };

    const handleCopySheet = (sheet) => {
        const data = { type: 'sheet', sourcePath: currentPath || project || '', sourceSheetId: sheet.name };
        setClipboard(data);
        setClipboardStorage(data);
        cancelPaste();
    };

    const clearClipboardFn = () => {
        setClipboard(null);
        setClipboardStorage(null);
        cancelPaste();
    };

    const startPaste = (target) => {
        // target: '__here__' for current folder, or subfolder name for paste-inside
        setPastingTarget(target || '__here__');
        if (!clipboard) return;
        if (clipboard.type === 'sheet') {
            setPasteName(clipboard.sourceSheetId || '');
        } else {
            const srcName = clipboard.sourcePath?.split('/').pop() || '';
            setPasteName(srcName ? srcName + '_copy' : '');
        }
    };

    const cancelPaste = () => {
        setPastingTarget(null);
        setPasteName('');
    };

    const confirmPaste = async () => {
        if (!clipboard) return;
        const name = pasteName.trim();
        if (!name) return;
        const base = currentPath || '';
        const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';

        try {
            let body;
            if (clipboard.type === 'sheet') {
                // Determine target folder
                let targetFolder;
                if (pastingTarget && pastingTarget !== '__here__') {
                    targetFolder = base ? `${base}/${pastingTarget}` : pastingTarget;
                } else {
                    targetFolder = base;
                }
                if (!targetFolder) {
                    alert('Cannot paste a sheet at the root level. Navigate into a project first.');
                    return;
                }
                body = {
                    source_type: 'sheet',
                    source_path: clipboard.sourcePath,
                    source_sheet_id: clipboard.sourceSheetId,
                    dest_path: targetFolder,
                    dest_name: name,
                };
            } else {
                // folder paste
                let destPath;
                if (pastingTarget && pastingTarget !== '__here__') {
                    destPath = base ? `${base}/${pastingTarget}/${name}` : `${pastingTarget}/${name}`;
                } else {
                    destPath = base ? `${base}/${name}` : name;
                }
                // Prevent pasting inside itself
                if (destPath === clipboard.sourcePath || destPath.startsWith(clipboard.sourcePath + '/')) {
                    alert('Cannot paste a folder inside itself.');
                    return;
                }
                body = {
                    source_type: 'folder',
                    source_path: clipboard.sourcePath,
                    dest_path: destPath,
                };
            }
            const res = await authenticatedFetch(`http://${host}:8082/api/projects/paste`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.ok) {
                cancelPaste();
                fetchFolders();
                fetchSheets();
            } else if (res.status === 409) {
                const text = await res.text();
                alert(text || 'Destination already exists');
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to paste');
            }
        } catch (e) {
            console.error('paste failed', e);
            alert('Unexpected error pasting folder');
        }
    };

    const toggleAuditSidebar = () => {
        if (isAuditOpen) {
            if (auditLogRef.current) {
                auditLogScrollTopRef.current = auditLogRef.current.scrollTop;
            }
            setAuditOpen(false);
        } else {
            setAuditOpen(true);
            fetchProjectAudit();
        }
    };

    const closeAuditSidebar = () => {
        if (auditLogRef.current) {
            auditLogScrollTopRef.current = auditLogRef.current.scrollTop;
        }
        setAuditOpen(false);
    };

    useEffect(() => {
        if (isAuditOpen && auditLogRef.current) {
            auditLogRef.current.scrollTop = auditLogScrollTopRef.current;
        }
    }, [isAuditOpen]);

    const deleteSheet = async (sheetId) => {
        if (!window.confirm('Are you sure you want to delete this sheet?')) {
            return;
        }

        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8082/api/sheets?id=${sheetId}${project ? `&project=${encodeURIComponent(project)}` : ''}` , {
                method: 'DELETE',
            });
            if (res.status === 403) {
                alert('Only the sheet owner can delete this sheet.');
                return;
            }
            if (res.ok) {
                fetchSheets();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                console.error('Failed to delete sheet');
                alert('Failed to delete sheet');
            }
        } catch (error) {
            console.error('Failed to delete sheet', error);
            alert('Error deleting sheet');
        }
    };

    const startRenaming = (sheet) => {
        setEditingSheetId(sheet.name);
        setEditingSheetName(sheet.name);
    };

    const cancelRenaming = () => {
        setEditingSheetId(null);
        setEditingSheetName('');
    };

    const renameSheet = async (sheetId) => {
        if (!editingSheetName.trim()) {
            alert('Sheet name cannot be empty');
            return;
        }

        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8082/api/sheets`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(currentPath ? { id: sheetId, name: editingSheetName, project_name: currentPath } : { id: sheetId, name: editingSheetName }),
            });
            if (res.status === 403) {
                alert('Only the sheet owner can rename this sheet.');
                return;
            }
            if (res.ok) {
                setEditingSheetId(null);
                setEditingSheetName('');
                fetchSheets();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                console.error('Failed to rename sheet');
                alert('Failed to rename sheet');
            }
        } catch (error) {
            console.error('Failed to rename sheet', error);
            alert('Error renaming sheet');
        }
    };

    // (Old startCopying/cancelCopying/copySheetToProject removed — unified into clipboard above)

    const displayedSheets = React.useMemo(() => {
        const list = sheets.slice();
        const q = searchQuery.trim().toLowerCase();
        if (!q) {
            return list.sort((a, b) => (a?.name || '').localeCompare(b?.name || ''));
        }
        return list.filter((s) => (s?.name || '').toLowerCase().includes(q));
    }, [sheets, searchQuery]);

    // Only the project owner may create/paste sheets and subfolders
    const isOwner = !projectOwner || projectOwner === username;

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
            {/* Bootstrap Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                     <button
                        onClick={goUpOne}
                        className="btn btn-outline-primary btn-sm d-flex align-items-center"
                    >
                        <ArrowLeft className="me-1" />
                    </button>
                    <a className="navbar-brand d-flex align-items-center" href="#">
                        <FileSpreadsheet className="me-2" />
                        {project ? `Project: ${project}` : 'SheetMaster'}
                    </a>
                    {project && (
                        <div className="d-flex align-items-center ms-auto">

                           
                            <button
                                onClick={handleDownloadProjectXlsx}
                                className="btn btn-outline-success btn-sm d-flex align-items-center me-2"
                                title="Export all sheets as XLSX"
                            >
                                <FileSpreadsheet className="me-1" /> Export XLSX
                            </button>
                            <input
                                ref={fileInputRef}
                                type="file"
                                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                style={{ display: 'none' }}
                                onChange={(e) => handleImportProjectXlsx(e.target.files?.[0])}
                            />
                            <button
                                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                                className="btn btn-outline-secondary btn-sm d-flex align-items-center me-2"
                                title="Import XLSX into project"
                            >
                                <FileSpreadsheet className="me-1" /> Import XLSX
                            </button>
                            <button
                                onClick={toggleAuditSidebar}
                                className={`btn btn-outline-primary btn-sm d-flex align-items-center me-2 ${isAuditOpen ? 'active' : ''}`}
                                title="Project Activity Log"
                            >
                                <History className="me-1" /> Activity
                            </button>
                            <span className="navbar-text me-3 d-flex align-items-center">
                                <User className="me-1" /> {username}
                            </span>
                            <button
                                onClick={handleLogout}
                                className="btn btn-outline-danger btn-sm d-flex align-items-center"
                                title="Logout"
                            >
                                <LogOut className="me-1" /> Logout
                            </button>
                        </div>
                    )}
                    {!project && (
                        <div className="d-flex align-items-center ms-auto">
                            <span className="navbar-text me-3 d-flex align-items-center">
                                <User className="me-1" /> {username}
                            </span>
                            <button
                                onClick={handleLogout}
                                className="btn btn-outline-danger btn-sm d-flex align-items-center"
                                title="Logout"
                            >
                                <LogOut className="me-1" /> Logout
                            </button>
                        </div>
                    )}
                    
                </div>
            </nav>

            <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
                {project && (
                    <div className="mb-3">
                        
                    {isAuditOpen && (
                        <div style={{ position: 'fixed', right: 16, top: 70, width: 360, height: 'calc(70% - 32px)', zIndex: 1100 }}>
                            <div className="d-flex justify-content-between align-items-center p-3 border-bottom bg-light">
                                <h5 className="mb-0 d-flex align-items-center">
                                    <History className="me-2" size={18} /> Project Activity
                                </h5>
                                <button 
                                    onClick={closeAuditSidebar} 
                                    className="btn btn-sm btn-light"
                                    aria-label="Close sidebar"
                                >
                                    ←
                                </button>
                            </div>
                            <div ref={auditLogRef} className="overflow-auto p-3" style={{ height: 'calc(70% - 56px)', overflowY: 'scroll' }}>
                                {auditLog.slice().reverse().map((entry, i) => {
                                    const ts = entry.timestamp ? new Date(entry.timestamp).toLocaleString() : '';
                                    const entryId = `${entry.timestamp || i}|${entry.user || ''}|${entry.action || ''}|${entry.details || ''}`;
                                    const isSelected = selectedAuditId === entryId;
                                    return (
                                        <div
                                            key={entryId}
                                            className={`p-2 mb-2 rounded ${isSelected ? 'bg-indigo-50' : 'bg-white'} border`}
                                            onClick={() => setSelectedAuditId(entryId)}
                                            title={ts}
                                        >
                                            <div className="d-flex justify-content-between">
                                                <span className="fw-semibold small">{entry.user}</span>
                                                <span className="text-muted small">{ts}</span>
                                            </div>
                                            <div className="small"><span className="badge bg-light text-dark me-2">{entry.action}</span>{entry.details}</div>
                                        </div>
                                    );
                                })}
                                {auditLog.length === 0 && (
                                    <div className="text-center text-muted py-5">
                                        <History className="mb-2" size={48} opacity={0.3} />
                                        <p className="mb-0">No project activity yet.</p>
                                    </div>
                                )}
                            </div>
                        </div>
                    )}
                    </div>
                )}
                
                 <div><h2>Sheets</h2></div>
                {/* Create Sheet - owner only */}
                {isOwner ? (
                <div className={`overflow-hidden transition-all duration-300 ease-in-out ${isCreating ? 'max-h-40 mb-8 opacity-100' : 'max-h-0 opacity-0'}`}>
                    <div className="p-6 bg-white border border-indigo-100 rounded-2xl shadow-sm">
                        <form onSubmit={createSheet} className="flex items-end gap-4">
                            <div className="flex-1">
                                <input
                                    type="text"
                                    value={newSheetName}
                                    onChange={(e) => setNewSheetName(e.target.value)}
                                    placeholder="New Sheet Name"
                                    className="w-full px-4 py-2 bg-gray-50 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 transition-all"
                                    autoFocus
                                />
                            <div className="d-flex align-items-center gap-2 mt-2">
                                <label className="form-label mb-0 text-sm">Type:</label>
                                <div className="d-flex gap-3">
                                    <label className="d-flex align-items-center gap-1 small" style={{ cursor: 'pointer' }}>
                                        <input
                                            type="radio"
                                            name="sheetType"
                                            value="datasheet"
                                            checked={newSheetType === 'datasheet'}
                                            onChange={() => setNewSheetType('datasheet')}
                                        />
                                        <span>📊 DataSheet</span>
                                        <span className="text-muted ms-1" style={{ fontSize: '0.75rem' }}>(flat, no hierarchy)</span>
                                    </label>
                                    <label className="d-flex align-items-center gap-1 small" style={{ cursor: 'pointer' }}>
                                        <input
                                            type="radio"
                                            name="sheetType"
                                            value="document"
                                            checked={newSheetType === 'document'}
                                            onChange={() => setNewSheetType('document')}
                                        />
                                        <span>📄 Document</span>
                                        <span className="text-muted ms-1" style={{ fontSize: '0.75rem' }}>(with row hierarchy)</span>
                                    </label>
                                </div>
                            </div>
                            <button
                                type="submit"
                                className="px-6 py-2 text-black font-medium rounded-full shadow-md transition-all hover:opacity-90 border-0 focus:outline-none mt-2"
                                style={{ backgroundColor: 'skyblue' }}
                            >
                                Create {newSheetType === 'document' ? 'Document' : 'DataSheet'}
                            </button>
                            </div>
                        </form>
                    </div>
                </div>
                ) : (
                <div className="mb-4 p-3 bg-white border border-warning rounded-2xl shadow-sm text-muted small">
                    Only the project owner can create sheets here.
                </div>
                )}
                {/* Actions Bar (Search only) */}
                <div className="flex flex-col md:flex-row justify-end items-start md:items-center gap-4 mb-8">
                    <div className="relative flex-1 md:w-64 group">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400 group-focus-within:text-indigo-500 transition-colors" />
                        <input
                            type="text"
                            placeholder="Search sheets..."
                            value={searchQuery}
                            onChange={(e) => setSearchQuery(e.target.value)}
                            className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm"
                        />
                    </div>
                </div>
                
                {/* List View Only - Table */}
                <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
                    <table className="table mb-0">
                        <thead className="table-group-header">
                            {/* give grey background to header */}
                            <tr style={{background: 'lightgray'}}>
                                <th scope="col">Sheet Name</th>
                                <th scope="col">Type</th>
                                <th scope="col">Owner Name</th>
                                <th scope="col" className="text-end">Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            {displayedSheets.map((sheet) => {
                                const sheetRoute = sheet.sheet_type === 'document' ? 'document' : 'sheet';
                                const sheetUrl = project ? `/${sheetRoute}/${sheet.name}?project=${encodeURIComponent(project)}` : `/${sheetRoute}/${sheet.name}`;
                                return (
                                <React.Fragment key={sheet.name}>
                                <tr style={{ cursor: 'pointer' }}>
                                    <td onClick={() => !editingSheetId && window.open(sheetUrl)}>
                                        {editingSheetId === sheet.name ? (
                                            <input
                                                type="text"
                                                className="form-control form-control-sm"
                                                value={editingSheetName}
                                                onChange={(e) => setEditingSheetName(e.target.value)}
                                                onClick={(e) => e.stopPropagation()}
                                                onKeyDown={(e) => {
                                                    if (e.key === 'Enter') {
                                                        renameSheet(sheet.name);
                                                    } else if (e.key === 'Escape') {
                                                        cancelRenaming();
                                                    }
                                                }}
                                                autoFocus
                                            />
                                        ) : (
                                            sheet.name
                                        )}
                                    </td>
                                    <td onClick={() => !editingSheetId && window.open(sheetUrl)}>
                                        {sheet.sheet_type === 'document'
                                            ? <span className="badge" style={{ backgroundColor: '#d4edda', color: '#155724' }}>📄 Document</span>
                                            : <span className="badge" style={{ backgroundColor: '#cce5ff', color: '#004085' }}>📊 DataSheet</span>
                                        }
                                    </td>
                                    <td onClick={() => !editingSheetId && window.open(sheetUrl)}>{sheet.owner}</td>
                                    <td className="text-end">
                                        {editingSheetId === sheet.name ? (
                                            <>
                                                <button
                                                    className="btn btn-sm btn-success me-2"
                                                    onClick={(ev) => { ev.stopPropagation(); renameSheet(sheet.name); }}
                                                >
                                                    Save
                                                </button>
                                                <button
                                                    className="btn btn-sm btn-secondary"
                                                    onClick={(ev) => { ev.stopPropagation(); cancelRenaming(); }}
                                                >
                                                    Cancel
                                                </button>
                                            </>
                                        ) : (
                                            <>
                                                {sheet.owner === username && (
                                                    <button
                                                        className="btn btn-sm btn-outline-primary me-2"
                                                        onClick={(ev) => { ev.stopPropagation(); startRenaming(sheet); }}
                                                    >
                                                        <Edit2 size={14} className="me-1" /> Rename
                                                    </button>
                                                )}
                                                <button
                                                    className="btn btn-sm btn-outline-primary me-2"
                                                    onClick={(ev) => { ev.stopPropagation(); handleCopySheet(sheet); }}
                                                >
                                                    <Copy size={14} className="me-1" /> Copy
                                                </button>
                                                {sheet.owner === username && (
                                                    <button
                                                        className="btn btn-sm btn-outline-danger"
                                                        onClick={(ev) => { ev.stopPropagation(); deleteSheet(sheet.name); }}
                                                    >
                                                        <Trash2 size={14} className="me-1" /> Delete
                                                    </button>
                                                )}
                                            </>
                                        )}
                                    </td>
                                </tr>
                                </React.Fragment>
                                );
                            })}
                            {displayedSheets.length === 0 && (
                                <tr>
                                    <td colSpan="4" className="text-center text-muted py-4">No sheets found.</td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>
                {/* Folders */}
                {currentPath && (
                    <div className="mb-6 mt-4">
                        <div><h2>SubFolders</h2></div>
                        {/* Clipboard banner — owner only for paste */}
                        {clipboard && isOwner && (
                            <div className="mb-3 p-2 bg-white border border-success rounded shadow-sm d-flex align-items-center gap-2 flex-wrap">
                                <span className="text-muted small">
                                    Clipboard ({clipboard.type === 'sheet' ? 'Sheet' : 'Folder'}): <strong>{clipboard.type === 'sheet' ? `${clipboard.sourcePath}/${clipboard.sourceSheetId}` : clipboard.sourcePath}</strong>
                                </span>
                                {pastingTarget === '__here__' ? (
                                    <>
                                        <input
                                            type="text"
                                            className="form-control form-control-sm"
                                            placeholder="Paste name"
                                            value={pasteName}
                                            onChange={(e) => setPasteName(e.target.value)}
                                            style={{ maxWidth: 220 }}
                                        />
                                        <button className="btn btn-sm btn-success" disabled={!pasteName.trim()} onClick={confirmPaste}>Paste</button>
                                        <button className="btn btn-sm btn-secondary" onClick={cancelPaste}>Cancel</button>
                                    </>
                                ) : (
                                    <button className="btn btn-sm btn-outline-success" onClick={() => startPaste('__here__')}>
                                        <ClipboardPaste size={14} className="me-1" /> Paste Here
                                    </button>
                                )}
                                <button className="btn btn-sm btn-outline-secondary ms-auto" onClick={clearClipboardFn}>
                                    <X size={14} className="me-1" /> Clear
                                </button>
                            </div>
                        )}
                        {clipboard && !isOwner && (
                            <div className="mb-3 p-2 bg-white border border-success rounded shadow-sm d-flex align-items-center gap-2 flex-wrap">
                                <span className="text-muted small">
                                    Clipboard ({clipboard.type === 'sheet' ? 'Sheet' : 'Folder'}): <strong>{clipboard.type === 'sheet' ? `${clipboard.sourcePath}/${clipboard.sourceSheetId}` : clipboard.sourcePath}</strong>
                                </span>
                                <span className="small text-warning ms-2">Only the project owner can paste here.</span>
                                <button className="btn btn-sm btn-outline-secondary ms-auto" onClick={clearClipboardFn}>
                                    <X size={14} className="me-1" /> Clear
                                </button>
                            </div>
                        )}

                        <div className="bg-white border rounded-2xl shadow-sm overflow-hidden">
                            {/* Create subfolder row — owner only */}
                            {isOwner && (
                                <div className="p-3 border-bottom d-flex align-items-center gap-2">
                                    <input
                                        type="text"
                                        className="form-control form-control-sm"
                                        value={newFolderName}
                                        onChange={(e) => setNewFolderName(e.target.value)}
                                        placeholder="New subfolder name"
                                        style={{ maxWidth: 280 }}
                                        onKeyDown={(e) => { if (e.key === 'Enter') createFolder(e); }}
                                    />
                                    <button className="btn btn-sm btn-primary" onClick={createFolder}>Create Folder</button>
                                </div>
                            )}
                            <table className="table table-hover mb-0 align-middle">
                                <thead style={{ background: 'lightgray' }}>
                                    <tr>
                                        <th>Folder Name</th>
                                        <th className="text-end">Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {folders.map((name) => {
                                        const fullPath = currentPath ? `${currentPath}/${name}` : name;
                                        const isClipboardSrc = clipboard && clipboard.type === 'folder' && clipboard.sourcePath === fullPath;
                                        return (
                                        <tr key={name}>
                                            <td style={{ cursor: 'pointer' }} onClick={() => editingFolderName !== name && goToFolder(name)}>
                                                {editingFolderName === name ? (
                                                    <input
                                                        type="text"
                                                        className="form-control form-control-sm"
                                                        value={editingFolderNewName}
                                                        onChange={(e) => setEditingFolderNewName(e.target.value)}
                                                        onClick={(e) => e.stopPropagation()}
                                                        onKeyDown={(e) => {
                                                            if (e.key === 'Enter') renameFolder(name);
                                                            if (e.key === 'Escape') cancelRenamingFolder();
                                                        }}
                                                        style={{ maxWidth: 220 }}
                                                        autoFocus
                                                    />
                                                ) : (
                                                    <span className="text-primary fw-semibold">{name}</span>
                                                )}
                                            </td>
                                            <td className="text-end">
                                                {editingFolderName === name ? (
                                                    <>
                                                        <button className="btn btn-sm btn-success me-1" onClick={() => renameFolder(name)}>Save</button>
                                                        <button className="btn btn-sm btn-secondary" onClick={cancelRenamingFolder}>Cancel</button>
                                                    </>
                                                ) : pastingTarget === name ? (
                                                    <div className="d-flex align-items-center gap-1 justify-content-end">
                                                        <span className="small text-muted">
                                                            Paste {clipboard?.type === 'sheet' ? 'sheet' : 'folder'} into <strong>{name}</strong>:
                                                        </span>
                                                        <input
                                                            type="text"
                                                            className="form-control form-control-sm"
                                                            placeholder="Paste name"
                                                            value={pasteName}
                                                            onChange={(e) => setPasteName(e.target.value)}
                                                            style={{ width: '140px' }}
                                                        />
                                                        <button className="btn btn-sm btn-success" disabled={!pasteName.trim()} onClick={confirmPaste}>Paste</button>
                                                        <button className="btn btn-sm btn-secondary" onClick={cancelPaste}>Cancel</button>
                                                    </div>
                                                ) : (
                                                    <div className="d-flex gap-1 justify-content-end">
                                                        {isOwner && (
                                                            <button className="btn btn-sm btn-outline-secondary" onClick={(e) => { e.stopPropagation(); startRenamingFolder(name); }} title="Rename folder">
                                                                <Edit2 size={14} />
                                                            </button>
                                                        )}
                                                        <button className={`btn btn-sm ${isClipboardSrc ? 'btn-primary' : 'btn-outline-secondary'}`} onClick={(e) => { e.stopPropagation(); handleCopyFolder(name); }} title="Copy folder">
                                                            <Copy size={14} />
                                                        </button>
                                                        {clipboard && !isClipboardSrc && isOwner && (
                                                            <button className="btn btn-sm btn-outline-success" onClick={(e) => { e.stopPropagation(); startPaste(name); }} title="Paste inside this folder">
                                                                <ClipboardPaste size={14} />
                                                            </button>
                                                        )}
                                                    </div>
                                                )}
                                            </td>
                                        </tr>
                                        );
                                    })}
                                    {folders.length === 0 && (
                                        <tr>
                                            <td colSpan={2} className="text-center text-muted py-3">No subfolders.</td>
                                        </tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}
            </main>
        </div>
    );
}
