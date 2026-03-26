import React, { useEffect, useState, useCallback } from 'react';
import { useNavigate, useParams, useSearchParams } from 'react-router-dom';
import { ArrowLeft, Plus, Save, X, Edit2, Trash2, Clock, User, LogOut, Calendar } from 'lucide-react';
import { isSessionValid, clearAuth, authenticatedFetch, getUsername, apiUrl } from '../utils/auth';

export default function Timeline() {
    const { project } = useParams();
    const navigate = useNavigate();
    const username = getUsername();

    const [entries, setEntries] = useState([]);
    const [loading, setLoading] = useState(true);
    const [projectOwner, setProjectOwner] = useState('');
    const [projectAdmins, setProjectAdmins] = useState([]);
    const isOwner = !projectOwner || projectOwner === username || projectAdmins.includes(username);

    // For new entry
    const [adding, setAdding] = useState(false);
    const [newTimestamp, setNewTimestamp] = useState('');
    const [newDescription, setNewDescription] = useState('');

    // For editing
    const [editingId, setEditingId] = useState(null);
    const [editTimestamp, setEditTimestamp] = useState('');
    const [editDescription, setEditDescription] = useState('');

    // Helper: convert ISO string from backend (time.Time → RFC3339) to datetime-local value
    const toDatetimeLocal = (isoStr) => {
        if (!isoStr) return '';
        try {
            const d = new Date(isoStr);
            if (isNaN(d.getTime())) return '';
            // Format as YYYY-MM-DDTHH:MM for datetime-local input
            const pad = n => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
        } catch { return ''; }
    };

    // Helper: format for display
    const formatTimestamp = (isoStr) => {
        if (!isoStr) return '';
        try {
            const d = new Date(isoStr);
            if (isNaN(d.getTime())) return isoStr;
            return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: '2-digit', hour: '2-digit', minute: '2-digit' });
        } catch { return isoStr; }
    };

    useEffect(() => {
        if (!username || !isSessionValid()) {
            clearAuth();
            navigate('/');
            return;
        }
        if (project) fetchTimeline();
        if (project) {
            authenticatedFetch(apiUrl('/api/projects'))
                .then(r => r.ok ? r.json() : [])
                .then(list => {
                    const topProject = project.split('/')[0];
                    const found = Array.isArray(list) ? list.find(p => p.name === topProject || p.id === topProject) : null;
                    setProjectOwner(found?.owner || '');
                    setProjectAdmins(Array.isArray(found?.admins) ? found.admins : []);
                })
                .catch(() => { setProjectOwner(''); setProjectAdmins([]); });
        }

        const interval = setInterval(() => {
            if (!isSessionValid()) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            }
        }, 60000);
        return () => clearInterval(interval);
    }, [project, username, navigate]);

    const fetchTimeline = useCallback(async () => {
        try {
            setLoading(true);
            const res = await authenticatedFetch(apiUrl(`/api/timeline?project=${encodeURIComponent(project)}`));
            if (res.ok) {
                const data = await res.json();
                setEntries(Array.isArray(data) ? data : []);
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                setEntries([]);
            }
        } catch (e) {
            console.error('Failed to fetch timeline', e);
            setEntries([]);
        } finally {
            setLoading(false);
        }
    }, [project, navigate]);

    const addEntry = async (e) => {
        e.preventDefault();
        if (!newTimestamp || !newDescription.trim()) return;
        try {
            const res = await authenticatedFetch(apiUrl('/api/timeline'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    project: project,
                    timestamp: newTimestamp,
                    description: newDescription.trim(),
                }),
            });
            if (res.ok) {
                setNewTimestamp('');
                setNewDescription('');
                setAdding(false);
                fetchTimeline();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to add entry');
            }
        } catch (err) {
            console.error('Failed to add timeline entry', err);
        }
    };

    const startEditing = (entry) => {
        setEditingId(entry.id);
        setEditTimestamp(toDatetimeLocal(entry.timestamp));
        setEditDescription(entry.description);
    };

    const cancelEditing = () => {
        setEditingId(null);
        setEditTimestamp('');
        setEditDescription('');
    };

    const saveEdit = async () => {
        if (!editTimestamp || !editDescription.trim()) return;
        try {
            const res = await authenticatedFetch(apiUrl('/api/timeline'), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    project: project,
                    id: editingId,
                    timestamp: editTimestamp,
                    description: editDescription.trim(),
                }),
            });
            if (res.ok) {
                cancelEditing();
                fetchTimeline();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to update entry');
            }
        } catch (err) {
            console.error('Failed to update timeline entry', err);
        }
    };

    const deleteEntry = async (id) => {
        if (!window.confirm('Are you sure you want to delete this timeline entry?')) return;
        try {
            const res = await authenticatedFetch(apiUrl(`/api/timeline?project=${encodeURIComponent(project)}&id=${encodeURIComponent(id)}`), {
                method: 'DELETE',
            });
            if (res.ok) {
                fetchTimeline();
            } else if (res.status === 401) {
                clearAuth();
                alert('Your session has expired. Please log in again.');
                navigate('/');
            } else {
                const text = await res.text();
                alert(text || 'Failed to delete entry');
            }
        } catch (err) {
            console.error('Failed to delete timeline entry', err);
        }
    };

    const handleLogout = async () => {
        try {
            await authenticatedFetch(apiUrl('/api/logout'), { method: 'POST' });
        } catch (e) {
            console.error('Logout error', e);
        } finally {
            clearAuth();
            navigate('/');
        }
    };

    const goBack = () => {
        if (project) {
            navigate(`/project/${encodeURIComponent(project)}`);
        } else {
            navigate('/projects');
        }
    };

    // Sort entries by timestamp (newest first)
    const sortedEntries = [...entries].sort((a, b) => {
        const tA = a.timestamp ? new Date(a.timestamp).getTime() : 0;
        const tB = b.timestamp ? new Date(b.timestamp).getTime() : 0;
        return tB - tA;
    });

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
            {/* Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                    <button onClick={goBack} className="btn btn-outline-primary btn-sm d-flex align-items-center">
                        <ArrowLeft className="me-1" />
                    </button>
                    <a className="navbar-brand d-flex align-items-center ms-2" href="#">
                        <Clock className="me-2" />
                        Timeline: {project}
                    </a>
                    <div className="d-flex align-items-center ms-auto">
                        <span className="navbar-text me-3 d-flex align-items-center">
                            <User className="me-1" /> {username}
                        </span>
                        <button onClick={handleLogout} className="btn btn-outline-danger btn-sm d-flex align-items-center" title="Logout">
                            <LogOut className="me-1" /> Logout
                        </button>
                    </div>
                </div>
            </nav>

            <main className="flex-1 max-w-4xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
                {/* Add Entry Button */}
                <div className="d-flex justify-content-between align-items-center mb-4">
                    <h2 className="mb-0 d-flex align-items-center">
                        <Calendar className="me-2" size={24} /> Project Timeline
                    </h2>
                    {!adding && isOwner && (
                        <button className="btn btn-primary d-flex align-items-center" onClick={() => {
                            setAdding(true);
                            // Pre-fill with current local datetime truncated to minutes
                            const now = new Date();
                            const pad = n => String(n).padStart(2, '0');
                            setNewTimestamp(`${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`);
                        }}>
                            <Plus className="me-1" size={16} /> Add Entry
                        </button>
                    )}
                </div>

                {/* Add Entry Form */}
                {adding && isOwner && (
                    <div className="card mb-4 shadow-sm">
                        <div className="card-body">
                            <h5 className="card-title mb-3">New Timeline Entry</h5>
                            <form onSubmit={addEntry}>
                                <div className="mb-3">
                                    <label className="form-label fw-semibold">Date &amp; Time</label>
                                    <input
                                        type="datetime-local"
                                        className="form-control"
                                        value={newTimestamp}
                                        onChange={(e) => setNewTimestamp(e.target.value)}
                                        required
                                    />
                                </div>
                                <div className="mb-3">
                                    <label className="form-label fw-semibold">Description</label>
                                    <textarea
                                        className="form-control"
                                        rows={3}
                                        value={newDescription}
                                        onChange={(e) => setNewDescription(e.target.value)}
                                        placeholder="Describe the event or milestone..."
                                        required
                                    />
                                </div>
                                <div className="d-flex gap-2">
                                    <button type="submit" className="btn btn-success d-flex align-items-center">
                                        <Save className="me-1" size={16} /> Save
                                    </button>
                                    <button type="button" className="btn btn-secondary d-flex align-items-center" onClick={() => { setAdding(false); setNewTimestamp(''); setNewDescription(''); }}>
                                        <X className="me-1" size={16} /> Cancel
                                    </button>
                                </div>
                            </form>
                        </div>
                    </div>
                )}

                {/* Loading */}
                {loading && (
                    <div className="text-center text-muted py-5">Loading timeline...</div>
                )}

                {/* Timeline Display */}
                {!loading && sortedEntries.length === 0 && !adding && (
                    <div className="text-center text-muted py-5">
                        <Clock className="mb-3" size={48} opacity={0.3} />
                        <p className="mb-0">No timeline entries yet. Click "Add Entry" to create one.</p>
                    </div>
                )}

                {!loading && sortedEntries.length > 0 && (
                    <div className="position-relative">
                        {/* Vertical line */}
                        <div
                            style={{
                                position: 'absolute',
                                left: 24,
                                top: 0,
                                bottom: 0,
                                width: 3,
                                backgroundColor: '#0d6efd',
                                borderRadius: 2,
                                zIndex: 0,
                            }}
                        />

                        {sortedEntries.map((entry, index) => (
                            <div key={entry.id} className="d-flex mb-4 position-relative" style={{ paddingLeft: 50 }}>
                                {/* Dot on timeline */}
                                <div
                                    style={{
                                        position: 'absolute',
                                        left: 16,
                                        top: 18,
                                        width: 18,
                                        height: 18,
                                        borderRadius: '50%',
                                        backgroundColor: '#0d6efd',
                                        border: '3px solid white',
                                        boxShadow: '0 0 0 2px #0d6efd',
                                        zIndex: 1,
                                    }}
                                />

                                <div className="card flex-grow-1 shadow-sm">
                                    <div className="card-body">
                                        {editingId === entry.id ? (
                                            /* Edit mode */
                                            <div>
                                                <div className="mb-2">
                                                    <label className="form-label fw-semibold small">Date &amp; Time</label>
                                                    <input
                                                        type="datetime-local"
                                                        className="form-control form-control-sm"
                                                        value={editTimestamp}
                                                        onChange={(e) => setEditTimestamp(e.target.value)}
                                                    />
                                                </div>
                                                <div className="mb-2">
                                                    <label className="form-label fw-semibold small">Description</label>
                                                    <textarea
                                                        className="form-control form-control-sm"
                                                        rows={3}
                                                        value={editDescription}
                                                        onChange={(e) => setEditDescription(e.target.value)}
                                                    />
                                                </div>
                                                <div className="d-flex gap-2">
                                                    <button className="btn btn-sm btn-success" onClick={saveEdit}>
                                                        <Save size={14} className="me-1" /> Save
                                                    </button>
                                                    <button className="btn btn-sm btn-secondary" onClick={cancelEditing}>
                                                        <X size={14} className="me-1" /> Cancel
                                                    </button>
                                                </div>
                                            </div>
                                        ) : (
                                            /* View mode */
                                            <div>
                                                <div className="d-flex justify-content-between align-items-start">
                                                    <div>
                                                        <span className="badge bg-primary me-2">
                                                            {formatTimestamp(entry.timestamp)}
                                                        </span>
                                                        <small className="text-muted">by {entry.user}</small>
                                                    </div>
                                                    <div className="d-flex gap-1">
                                                        {isOwner && (
                                                        <button
                                                            className="btn btn-sm btn-outline-primary"
                                                            onClick={() => startEditing(entry)}
                                                            title="Edit"
                                                        >
                                                            <Edit2 size={14} />
                                                        </button>
                                                        )}
                                                        {isOwner && (
                                                        <button
                                                            className="btn btn-sm btn-outline-danger"
                                                            onClick={() => deleteEntry(entry.id)}
                                                            title="Delete"
                                                        >
                                                            <Trash2 size={14} />
                                                        </button>
                                                        )}
                                                    </div>
                                                </div>
                                                <p className="mt-2 mb-0" style={{ whiteSpace: 'pre-wrap' }}>{entry.description}</p>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </div>
                        ))}
                    </div>
                )}
            </main>
        </div>
    );
}
