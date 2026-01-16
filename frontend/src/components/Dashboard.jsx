import React, { useEffect, useState } from 'react';
import { useNavigate, useParams, Link } from 'react-router-dom';
import {
    FileSpreadsheet,
    Plus,
    Search,
    LogOut,
    User,
    MoreVertical,
    Trash2,
    Edit2
} from 'lucide-react';
import './bootstrap/dist/css/bootstrap.min.css';
import { isSessionValid, clearAuth, authenticatedFetch, getUsername } from '../utils/auth';

export default function Dashboard() {
    const { project } = useParams();
    const [sheets, setSheets] = useState([]);
    const [newSheetName, setNewSheetName] = useState('');
    const [isCreating, setIsCreating] = useState(true);
    const [searchQuery, setSearchQuery] = useState('');
    const [editingSheetId, setEditingSheetId] = useState(null);
    const [editingSheetName, setEditingSheetName] = useState('');
    const navigate = useNavigate();
    const username = getUsername();

    useEffect(() => {
        // Check session validity on mount and periodically
        if (!username || !isSessionValid()) {
            clearAuth();
            navigate('/');
            return;
        }

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
                navigate(`/sheet/${sheet.id}`);
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
            await authenticatedFetch('http://localhost:8080/api/logout', {
                method: 'POST',
            });
        } catch (error) {
            console.error("Logout error", error);
        } finally {
            clearAuth();
            navigate('/');
        }
    };

    const deleteSheet = async (sheetId) => {
        if (!window.confirm('Are you sure you want to delete this sheet?')) {
            return;
        }

        try {
            const res = await authenticatedFetch(`http://localhost:8080/api/sheets?id=${sheetId}`, {
                method: 'DELETE',
            });
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
            const res = await authenticatedFetch('http://localhost:8080/api/sheets', {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: sheetId, name: editingSheetName }),
            });
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
                    <a className="navbar-brand d-flex align-items-center" href="#">
                        <FileSpreadsheet className="me-2" />
                        {project ? `Project: ${project}` : 'SheetMaster'}
                    </a>
                    <button
                        className="navbar-toggler"
                        type="button"
                        data-toggle="collapse"
                        data-target="#mainNavbar"
                        aria-controls="mainNavbar"
                        aria-expanded="false"
                        aria-label="Toggle navigation"
                    >
                        <span className="navbar-toggler-icon"></span>
                    </button>
                    <div className="collapse navbar-collapse" id="mainNavbar">
                        <ul className="navbar-nav mr-auto">
                            <li className="nav-item">
                                
                            </li>
                        </ul>
                        <div className="d-flex align-items-center">
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
                    </div>
                </div>
            </nav>

            <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
                {project && (
                    <div className="mb-3">
                        <Link to="/projects" className="btn btn-sm btn-outline-secondary">← Back to Projects</Link>
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
                                <tr key={sheet.id} style={{ cursor: 'pointer' }}>
                                    <td onClick={() => !editingSheetId &&  window.open(`/sheet/${sheet.id}`)}>
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
                                    <td onClick={() => !editingSheetId &&  window.open(`/sheet/${sheet.id}`)}>{sheet.owner}</td>
                                    <td className="font-mono" onClick={() => !editingSheetId &&  window.open(`/sheet/${sheet.id}`)}>{sheet.id}</td>
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
                                                <button
                                                    className="btn btn-sm btn-outline-primary me-2"
                                                    onClick={(ev) => { ev.stopPropagation(); startRenaming(sheet); }}
                                                >
                                                    <Edit2 size={14} className="me-1" /> Rename
                                                </button>
                                                <button
                                                    className="btn btn-sm btn-outline-danger"
                                                    onClick={(ev) => { ev.stopPropagation(); deleteSheet(sheet.id); }}
                                                >
                                                    <Trash2 size={14} className="me-1" /> Delete
                                                </button>
                                            </>
                                        )}
                                    </td>
                                </tr>
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
