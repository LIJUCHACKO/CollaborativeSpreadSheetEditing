import React, { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { authenticatedFetch, isSessionValid, clearAuth, getUsername, apiUrl, isAdmin, canCreateProject } from '../utils/auth';
import { Copy, ClipboardPaste, Edit2, Trash2, Search, User, LogOut, Folder, Lock, X, ShieldCheck } from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';

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

export default function Projects() {
  const [projects, setProjects] = useState([]);
  const [search, setSearch] = useState('');
  const [newProject, setNewProject] = useState('');
  const [editingName, setEditingName] = useState('');
  const [editingProject, setEditingProject] = useState(null);
  const [deleteProject, setDeleteProject] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState('');
  const [clipboard, setClipboard] = useState(() => getClipboard()); // {type:'folder'|'sheet', sourcePath, sourceSheetId?}
  const [pastingTarget, setPastingTarget] = useState(null); // project name or '__project_root__'
  const [pasteName, setPasteName] = useState('');
  // ...existing code...
  const navigate = useNavigate();
  const username = getUsername();

  useEffect(() => {
    if (!username || !isSessionValid()) {
      clearAuth();
      navigate('/');
      return;
    }
    const interval = setInterval(() => {
      if (!isSessionValid()) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    }, 60000);
    fetchProjects();
    return () => clearInterval(interval);
  }, [username, navigate]);

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

  const fetchProjects = async () => {
    try {
      const res = await authenticatedFetch(apiUrl('/api/projects'));
      if (res.ok) {
        const data = await res.json();
        setProjects(Array.isArray(data) ? data : []);
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    } catch (e) {
      console.error('fetch projects failed', e);
    }
  };

  const formatTimestamp = (ts) => {
    if (!ts) return '';
    try {
      const d = new Date(ts);
      if (isNaN(d.getTime())) return ts;
      return d.toLocaleString();
    } catch {
      return ts;
    }
  };

  const createProject = async (e) => {
    e.preventDefault();
    const name = newProject.trim();
    if (!name) return;
    try {
      const res = await authenticatedFetch(apiUrl('/api/projects'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name }),
      });
      if (res.ok) {
        setNewProject('');
        fetchProjects();
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    } catch (e) {
      console.error('create project failed', e);
    }
  };

  const startRename = (projectName) => {
    setEditingProject(projectName);
    setEditingName(projectName || '');
  };

  const cancelRename = () => {
    setEditingProject(null);
    setEditingName('');
  };

  const renameProject = async () => {
    const oldName = editingProject;
    const newName = editingName.trim();
    if (!oldName || !newName) return;
    try {
      const res = await authenticatedFetch(apiUrl('/api/projects'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ OldName: oldName, NewName: newName }),
      });
      if (res.status === 403) {
        alert('Only the project owner can rename this project.');
        return;
      }
      if (res.status === 409) {
        const errorText = await res.text();
        alert(errorText || 'Cannot rename: one or more sheets in this project are currently open by users.');
        return;
      }
      if (res.ok) {
        setEditingProject(null);
        setEditingName('');
        fetchProjects();
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    } catch (e) {
      console.error('rename project failed', e);
    }
  };

  const requestDelete = (projectName) => {
    setDeleteProject(projectName);
    setDeleteConfirm('');
  };

  const cancelDelete = () => {
    setDeleteProject(null);
    setDeleteConfirm('');
  };

  const handleCopy = (projectName) => {
    const data = { type: 'folder', sourcePath: projectName };
    setClipboard(data);
    setClipboardStorage(data);
    // Clear any paste UI that might be open
    setPastingTarget(null);
    setPasteName('');
  };

  const clearClipboard = () => {
    setClipboard(null);
    setClipboardStorage(null);
    cancelPaste();
  };

  const startPaste = (targetContext) => {
    // targetContext: '__project_root__' for top-level, or project name for inside that project
    setPastingTarget(targetContext || '__project_root__');
    const srcName = clipboard?.sourcePath?.split('/').pop() || '';
    setPasteName(srcName ? srcName + '_copy' : '');
  };

  const cancelPaste = () => {
    setPastingTarget(null);
    setPasteName('');
  };

  const confirmPaste = async () => {
    if (!clipboard) return;
    const name = pasteName.trim();
    if (!name) return;

    let destPath;
    if (pastingTarget === '__project_root__') {
      // Paste as new top-level project
      destPath = name;
    } else {
      // Paste inside an existing project
      destPath = pastingTarget + '/' + name;
    }

    const source = clipboard.sourcePath;
    // For folders: prevent pasting inside itself
    if (clipboard.type === 'folder') {
      if (destPath === source || destPath.startsWith(source + '/')) {
        alert('Cannot paste a folder inside itself.');
        return;
      }
    }

    try {
      const body = clipboard.type === 'sheet'
        ? { source_type: 'sheet', source_path: clipboard.sourcePath, source_sheet_id: clipboard.sourceSheetId, dest_path: pastingTarget === '__project_root__' ? name : pastingTarget, dest_name: name }
        : { source_type: 'folder', source_path: source, dest_path: destPath };

      const res = await authenticatedFetch(apiUrl('/api/projects/paste'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (res.ok) {
        cancelPaste();
        fetchProjects();
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
      alert('Unexpected error pasting');
    }
  };

  // ...existing code...

  const confirmDelete = async () => {
    const name = deleteProject;
    if (!name) return;
    if (deleteConfirm.trim() !== name) return; // require exact match
    try {
      const res = await authenticatedFetch(apiUrl(`/api/projects?name=${encodeURIComponent(name)}`), {
        method: 'DELETE',
      });
      if (res.status === 403) {
        alert('Only the project owner can delete this project.');
        return;
      }
      if (res.ok) {
        cancelDelete();
        fetchProjects();
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    } catch (e) {
      console.error('delete project failed', e);
    }
  };

  const displayed = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return [...projects].sort((a,b) => (a?.name||'').localeCompare(b?.name||''));
    return projects.filter(p => (p?.name||'').toLowerCase().includes(q));
  }, [projects, search]);

  const handleLogout = async () => {
    try {
      await authenticatedFetch(apiUrl('/api/logout'), { method: 'POST' });
    } catch {}
    clearAuth();
    navigate('/');
  };

  // Password change moved to dedicated page

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
      <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
        <div className="container-fluid">
          <a className="navbar-brand d-flex align-items-center" href="#">
            <Folder className="me-2" /> Projects
          </a>
          <div className="d-flex align-items-center">
            <span className="navbar-text me-3 d-flex align-items-center">
              <User className="me-1" /> {username}
            </span>
            {isAdmin() && (
              <button onClick={() => navigate('/admin')} className="btn btn-outline-dark btn-sm d-flex align-items-center me-2" title="Admin Panel">
                <ShieldCheck size={14} className="me-1" /> Admin
              </button>
            )}
            <button onClick={() => navigate('/change-password')} className="btn btn-outline-primary btn-sm d-flex align-items-center me-2" title="Change Password">
              <Lock className="me-1" /> Change Password
            </button>
            <button onClick={handleLogout} className="btn btn-outline-danger btn-sm d-flex align-items-center" title="Logout">
              <LogOut className="me-1" /> Logout
            </button>
          </div>
        </div>
      </nav>

      <main className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        

        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-8">
          <div className="relative flex-1 md:w-64 group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400 group-focus-within:text-indigo-500 transition-colors" />
            <input type="text" placeholder="Search projects..." value={search} onChange={(e)=>setSearch(e.target.value)} className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm" />
          </div>
          <div className={`overflow-hidden transition-all duration-300 ease-in-out  max-h-40 mb-8 opacity-100 }`}>
          {canCreateProject() ? (
          <div className="p-6 bg-white border border-indigo-100 rounded-2xl shadow-sm">
            <form onSubmit={createProject} className="flex items-end gap-4">
              <div className="flex-1">
                <input type="text" value={newProject} onChange={(e)=>setNewProject(e.target.value)} placeholder="New Project Name" className="w-full px-4 py-2 bg-gray-50 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 transition-all" autoFocus />
                <div className="mt-2 d-inline-flex gap-2">
                  <button type="submit" className="px-6 py-2 text-black font-medium rounded-full shadow-md transition-all hover:opacity-90 border-0 focus:outline-none" style={{ backgroundColor: 'skyblue' }}>Create</button>
                  <button type="button" className="btn btn-sm btn-secondary" onClick={() => { setNewProject(''); }}>Cancel</button>
                </div>
              </div>
            </form>
          </div>
          ) : (
          <div className="p-3 bg-white border border-warning rounded-2xl shadow-sm text-muted small d-flex align-items-center gap-2">
            <ShieldCheck size={16} className="text-warning" />
            You are not allowed to create projects. Contact an admin to request permission.
          </div>
          )}
        </div>
        </div>

        {clipboard && canCreateProject() && (
          <div className="mb-4 p-3 bg-white border border-success rounded-2xl shadow-sm d-flex align-items-center gap-3 flex-wrap">
            <span className="text-muted small">
              Clipboard: <strong>{clipboard.type === 'sheet' ? `Sheet: ${clipboard.sourceSheetId}` : clipboard.sourcePath}</strong>
              {clipboard.type === 'sheet' && clipboard.sourcePath && <span className="ms-1">(from {clipboard.sourcePath})</span>}
            </span>
            {pastingTarget === '__project_root__' ? (
              <>
                <input type="text" className="form-control form-control-sm" placeholder="New project name" value={pasteName} onChange={(e)=>setPasteName(e.target.value)} style={{ maxWidth: 220 }} />
                <button className="btn btn-sm btn-success" disabled={!pasteName.trim()} onClick={confirmPaste}>Paste</button>
                <button className="btn btn-sm btn-secondary" onClick={cancelPaste}>Cancel</button>
              </>
            ) : (
              clipboard.type === 'folder' && (
                <button className="btn btn-sm btn-outline-success" onClick={() => startPaste('__project_root__')}>
                  <ClipboardPaste size={14} className="me-1" /> Paste as New Project
                </button>
              )
            )}
            <button className="btn btn-sm btn-outline-secondary ms-auto" onClick={clearClipboard}>
              <X size={14} className="me-1" /> Clear
            </button>
          </div>
        )}

        <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
          <table className="table mb-0">
            <thead>
              <tr style={{background: 'lightgray'}}>
                <th scope="col">Project</th>
                <th scope="col">Owner</th>
                <th scope="col" className="text-end">Actions</th>
              </tr>
            </thead>
            <tbody>
              {displayed.map((p, idx) => (
                <tr key={p.name} style={{ cursor: 'pointer' }}>
                  <td onClick={() => navigate(`/project/${encodeURIComponent(p.name)}`)}>
                    {editingProject === p.name ? (
                      <input type="text" className="form-control form-control-sm" value={editingName} onChange={(e)=>setEditingName(e.target.value)} onClick={(e)=>e.stopPropagation()} onKeyDown={(e)=>{ if (e.key==='Enter') renameProject(); if (e.key==='Escape') cancelRename(); }} autoFocus />
                    ) : p.name}
                  </td>
                  <td>
                    {p.owner || ''}
                  </td>
                  <td className="text-end">
                    {editingProject === p.name ? (
                      <>
                        <button className="btn btn-sm btn-success me-2" onClick={(ev)=>{ev.stopPropagation(); renameProject();}}>Save</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelRename();}}>Cancel</button>
                      </>
                    ) : deleteProject === p.name ? (
                      <div className="d-inline-flex align-items-center gap-2">
                        <input type="text" className="form-control form-control-sm" placeholder={`Type '${p.name}'`} value={deleteConfirm} onChange={(e)=>setDeleteConfirm(e.target.value)} style={{ maxWidth: 200 }} />
                        <button className="btn btn-sm btn-danger" disabled={deleteConfirm.trim()!==p.name} onClick={(ev)=>{ev.stopPropagation(); confirmDelete();}}>Confirm</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelDelete();}}>Cancel</button>
                      </div>
                    ) : pastingTarget === p.name ? (
                      <div className="d-inline-flex align-items-center gap-2">
                        <span className="text-muted small me-1">Paste into <strong>{p.name}/</strong>:</span>
                        <input type="text" className="form-control form-control-sm" placeholder={`Name inside ${p.name}`} value={pasteName} onChange={(e)=>setPasteName(e.target.value)} style={{ maxWidth: 220 }} onClick={(e)=>e.stopPropagation()} />
                        <button className="btn btn-sm btn-success" disabled={!pasteName.trim()} onClick={(ev)=>{ev.stopPropagation(); confirmPaste();}}>Paste</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelPaste();}}>Cancel</button>
                      </div>
                    ) : (
                      <>
                        {p.owner === username && (
                          <button className="btn btn-sm btn-outline-primary me-2" onClick={(ev)=>{ev.stopPropagation(); startRename(p.name);}}>
                            <Edit2 size={14} className="me-1"/> Rename
                          </button>
                        )}
                        {canCreateProject() && (
                          <button className={`btn btn-sm me-2 ${clipboard?.sourcePath === p.name && clipboard?.type === 'folder' ? 'btn-primary' : 'btn-outline-secondary'}`} onClick={(ev)=>{ev.stopPropagation(); handleCopy(p.name);}}>
                            <Copy size={14} className="me-1"/> {clipboard?.sourcePath === p.name && clipboard?.type === 'folder' ? 'Copied' : 'Copy'}
                          </button>
                        )}
                        {canCreateProject() && clipboard && !(clipboard.type === 'folder' && clipboard.sourcePath === p.name) && (
                          <button className="btn btn-sm btn-outline-success me-2" onClick={(ev)=>{ev.stopPropagation(); startPaste(p.name);}} title={`Paste inside ${p.name}`}>
                            <ClipboardPaste size={14} className="me-1"/> Paste Inside
                          </button>
                        )}
                        {p.owner === username && (
                          <button className="btn btn-sm btn-outline-danger" onClick={(ev)=>{ev.stopPropagation(); requestDelete(p.name);}}>
                            <Trash2 size={14} className="me-1"/> Delete
                          </button>
                        )}
                      </>
                    )}
                  </td>
                </tr>
              ))}
              {displayed.length === 0 && (
                <tr><td colSpan="3" className="text-center text-muted py-4">No projects found.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </main>
      {/* Audit modal removed */}
    </div>
  );
}
