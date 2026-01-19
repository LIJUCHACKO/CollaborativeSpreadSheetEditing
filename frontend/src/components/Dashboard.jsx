import React, { useEffect, useState, useRef } from 'react';
import { useNavigate, useParams, Link } from 'react-router-dom';
import {
    FileSpreadsheet,
    Plus,
    Search,
    LogOut,
    User,
    MoreVertical,
    Trash2,
    Edit2,
    History,
    ArrowLeft
} from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';
import { isSessionValid, clearAuth, authenticatedFetch, getUsername } from '../utils/auth';

export default function Dashboard() {
    const { project } = useParams();
    const [sheets, setSheets] = useState([]);
    const [newSheetName, setNewSheetName] = useState('');
    const [isCreating, setIsCreating] = useState(true);
    const [searchQuery, setSearchQuery] = useState('');
    const [editingSheetId, setEditingSheetId] = useState(null);
    const [editingSheetName, setEditingSheetName] = useState('');
    const [copyingSheetId, setCopyingSheetId] = useState(null);
    const [copyName, setCopyName] = useState('');
    const [targetProject, setTargetProject] = useState('');
    const [projectsList, setProjectsList] = useState([]);
    const navigate = useNavigate();
    const username = getUsername();
    // Project audit sidebar state
    const [auditLog, setAuditLog] = useState([]);
    const [isAuditOpen, setAuditOpen] = useState(false);
    const [selectedAuditId, setSelectedAuditId] = useState(null);
    const auditLogRef = useRef(null);
    const auditLogScrollTopRef = useRef(0);
    const fileInputRef = useRef(null);

    // Download all sheets in project as XLSX
    const handleDownloadProjectXlsx = async () => {
        try {
            if (!project) return;
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8080/api/export_project?project=${encodeURIComponent(project)}`, {
                method: 'GET',
            });
            if (!res.ok) {
                const text = await res.text();
                alert(`Failed to export project: ${text}`);
                return;
            }
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const safeName = (project || 'project') + '.xlsx';
            a.download = safeName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Error downloading project XLSX', err);
            alert('An unexpected error occurred while exporting the project.');
        }
    };

    // Import XLSX into current project (creates sheets per workbook sheet)
    const handleImportProjectXlsx = async (file) => {
        try {
            if (!project) return;
            if (!file) return;
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const form = new FormData();
            form.append('file', file);
            const res = await authenticatedFetch(`http://${host}:8080/api/import_project_xlsx?project=${encodeURIComponent(project)}`, {
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

    useEffect(() => {
        // Check session validity on mount and periodically
        if (!username || !isSessionValid()) {
            clearAuth();
            navigate('/');
            return;
        }
        // Load project audit entries
        fetchProjectAudit();

        // Check session every minute
        const sessionCheckInterval = setInterval(() => {
            if (!isSessionValid()) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        }, 60000); // Check every minute

        fetchSheets();

        return () => clearInterval(sessionCheckInterval);
    }, [username, navigate]);

    const fetchSheets = async () => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const query = project ? `?project=${encodeURIComponent(project)}` : '';
            const res = await authenticatedFetch(`http://${host}:8080/api/sheets${query}`);
            if (res.ok) {
                const data = await res.json();
                setSheets(data || []);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        } catch (error) {
            console.error("Failed to fetch sheets", error);
        }
    };

    const createSheet = async (e) => {
        e.preventDefault();
        if (!newSheetName.trim()) return;

        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const body = project ? { name: newSheetName, user: username, project_name: project } : { name: newSheetName, user: username };
            const res = await authenticatedFetch(`http://${host}:8080/api/sheets`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.ok) {
                const sheet = await res.json();
                setNewSheetName('');
                setIsCreating(false);
                fetchSheets();
                navigate(project ? `/sheet/${sheet.id}?project=${encodeURIComponent(project)}` : `/sheet/${sheet.id}`);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        } catch (error) {
            console.error("Failed to create sheet", error);
        }
    };

    const handleLogout = async () => {
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            await authenticatedFetch(`http://${host}:8080/api/logout`, {
                method: 'POST',
            });
        } catch (error) {
            console.error("Logout error", error);
        } finally {
            clearAuth();
            navigate('/');
        }
    };

    // Fetch project audit entries
    const fetchProjectAudit = async () => {
        try {
            if (!project) { setAuditLog([]); return; }
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8080/api/projects/audit?project=${encodeURIComponent(project)}`);
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
            const res = await authenticatedFetch(`http://${host}:8080/api/sheets?id=${sheetId}${project ? `&project=${encodeURIComponent(project)}` : ''}` , {
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
        setEditingSheetId(sheet.id);
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
            const res = await authenticatedFetch(`http://${host}:8080/api/sheets`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(project ? { id: sheetId, name: editingSheetName, project_name: project } : { id: sheetId, name: editingSheetName }),
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

    const startCopying = async (sheet) => {
        setCopyingSheetId(sheet.id);
        setCopyName(sheet.name ? `${sheet.name} (Copy)` : 'Copy');
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const res = await authenticatedFetch(`http://${host}:8080/api/projects`);
            if (res.ok) {
                const list = await res.json();
                const names = Array.isArray(list) ? list.map(p => p.name) : [];
                setProjectsList(names);
                // Preselect different project if current project exists
                if (project && names.length > 0) {
                    const alt = names.find(n => n !== project) || names[0];
                    setTargetProject(alt);
                } else if (names.length > 0) {
                    setTargetProject(names[0]);
                }
            }
        } catch (e) {
            // ignore fetch error
        }
    };

    const cancelCopying = () => {
        setCopyingSheetId(null);
        setCopyName('');
        setTargetProject('');
    };

    const copySheetToProject = async (sheet) => {
        const sourceProject = project || sheet.project_name || '';
        if (!targetProject) {
            alert('Select target project');
            return;
        }
        try {
            const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
            const body = {
                source_id: sheet.id,
                source_project: sourceProject,
                target_project: targetProject,
                name: copyName || sheet.name,
            };
            const res = await authenticatedFetch(`http://${host}:8080/api/sheet/copy`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (res.ok) {
                const newSheet = await res.json();
                cancelCopying();
                // If current list is filtered by project and target is same, refresh
                fetchSheets();
                // Navigate to new sheet in target project
                if (newSheet?.id) {
                    const destProject = targetProject;
                    window.open(`/sheet/${newSheet.id}?project=${encodeURIComponent(destProject)}`);
                }
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to copy sheet');
            }
        } catch (e) {
            console.error('copy sheet failed', e);
            alert('Unexpected error copying sheet');
        }
    };

    const displayedSheets = React.useMemo(() => {
        const list = sheets.slice();
        const q = searchQuery.trim().toLowerCase();
        if (!q) {
            return list.sort((a, b) => (a?.name || '').localeCompare(b?.name || ''));
        }
        return list.filter((s) => (s?.name || '').toLowerCase().includes(q));
    }, [sheets, searchQuery]);

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
            {/* Bootstrap Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                     <button
                        onClick={() => {
                            
                                navigate('/projects');
                            
                        }}
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
                

                {/* Create Sheet Modal/Collapse */}
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
                            
                            <button
                                type="submit"
                                className="px-6 py-2 text-black font-medium rounded-full shadow-md transition-all hover:opacity-90 border-0 focus:outline-none"
                                style={{ backgroundColor: 'skyblue' }}
                            >
                                Create
                            </button>
                            </div>
                        </form>
                    </div>
                </div>
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
                                <th scope="col">Owner Name</th>
                                <th scope="col">ID</th>
                                <th scope="col" className="text-end">Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            {displayedSheets.map((sheet) => (
                                <React.Fragment key={sheet.id}>
                                <tr style={{ cursor: 'pointer' }}>
                                    <td onClick={() => !editingSheetId &&  window.open(project ? `/sheet/${sheet.id}?project=${encodeURIComponent(project)}` : `/sheet/${sheet.id}`)}>
                                        {editingSheetId === sheet.id ? (
                                            <input
                                                type="text"
                                                className="form-control form-control-sm"
                                                value={editingSheetName}
                                                onChange={(e) => setEditingSheetName(e.target.value)}
                                                onClick={(e) => e.stopPropagation()}
                                                onKeyDown={(e) => {
                                                    if (e.key === 'Enter') {
                                                        renameSheet(sheet.id);
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
                                    <td onClick={() => !editingSheetId &&  window.open(project ? `/sheet/${sheet.id}?project=${encodeURIComponent(project)}` : `/sheet/${sheet.id}`)}>{sheet.owner}</td>
                                    <td className="font-mono" onClick={() => !editingSheetId &&  window.open(project ? `/sheet/${sheet.id}?project=${encodeURIComponent(project)}` : `/sheet/${sheet.id}`)}>{sheet.id}</td>
                                    <td className="text-end">
                                        {editingSheetId === sheet.id ? (
                                            <>
                                                <button
                                                    className="btn btn-sm btn-success me-2"
                                                    onClick={(ev) => { ev.stopPropagation(); renameSheet(sheet.id); }}
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
                                                    onClick={(ev) => { ev.stopPropagation(); startCopying(sheet); }}
                                                >
                                                    <Plus size={14} className="me-1" /> Copy
                                                </button>
                                                {sheet.owner === username && (
                                                    <button
                                                        className="btn btn-sm btn-outline-danger"
                                                        onClick={(ev) => { ev.stopPropagation(); deleteSheet(sheet.id); }}
                                                    >
                                                        <Trash2 size={14} className="me-1" /> Delete
                                                    </button>
                                                )}
                                            </>
                                        )}
                                    </td>
                                </tr>
                                {copyingSheetId === sheet.id && (
                                    <tr>
                                        <td colSpan="4">
                                            <div className="d-flex align-items-center gap-2">
                                                <select className="form-select form-select-sm" value={targetProject} onChange={(e)=>setTargetProject(e.target.value)} style={{ maxWidth: 220 }}>
                                                    <option value="">Select target project</option>
                                                    {projectsList.map((pname)=> (
                                                        <option key={pname} value={pname}>{pname}</option>
                                                    ))}
                                                </select>
                                                <input type="text" className="form-control form-control-sm" value={copyName} onChange={(e)=>setCopyName(e.target.value)} placeholder="Copy name" style={{ maxWidth: 260 }} />
                                                <button className="btn btn-sm btn-success" onClick={(ev)=>{ ev.stopPropagation(); copySheetToProject(sheet); }}>Copy Here</button>
                                                <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ ev.stopPropagation(); cancelCopying(); }}>Cancel</button>
                                            </div>
                                        </td>
                                    </tr>
                                )}
                                </React.Fragment>
                            ))}
                            {displayedSheets.length === 0 && (
                                <tr>
                                    <td colSpan="4" className="text-center text-muted py-4">No sheets found.</td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>
            </main>
        </div>
    );
}
