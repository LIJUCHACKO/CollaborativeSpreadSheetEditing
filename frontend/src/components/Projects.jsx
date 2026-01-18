import React, { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { authenticatedFetch, isSessionValid, clearAuth, getUsername } from '../utils/auth';
import { FolderPlus, Edit2, Trash2, Search, User, LogOut, Folder } from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';

export default function Projects() {
  const [projects, setProjects] = useState([]);
  const [search, setSearch] = useState('');
  const [creating, setCreating] = useState(false);
  const [newProject, setNewProject] = useState('');
  const [editingName, setEditingName] = useState('');
  const [editingIdx, setEditingIdx] = useState(null);
  const [deleteIdx, setDeleteIdx] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState('');
  const [duplicateIdx, setDuplicateIdx] = useState(null);
  const [duplicateName, setDuplicateName] = useState('');
  const [auditOpen, setAuditOpen] = useState(false);
  const [auditProject, setAuditProject] = useState('');
  const [auditEntries, setAuditEntries] = useState([]);
  const [auditLoading, setAuditLoading] = useState(false);
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

  const fetchProjects = async () => {
    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects`);
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
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name }),
      });
      if (res.ok) {
        setNewProject('');
        setCreating(false);
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

  const startRename = (idx) => {
    setEditingIdx(idx);
    setEditingName(projects[idx]?.name || '');
  };

  const cancelRename = () => {
    setEditingIdx(null);
    setEditingName('');
  };

  const renameProject = async (idx) => {
    const oldName = projects[idx]?.name;
    const newName = editingName.trim();
    if (!oldName || !newName) return;
    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ OldName: oldName, NewName: newName }),
      });
      if (res.ok) {
        setEditingIdx(null);
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

  const requestDelete = (idx) => {
    setDeleteIdx(idx);
    setDeleteConfirm('');
  };

  const cancelDelete = () => {
    setDeleteIdx(null);
    setDeleteConfirm('');
  };

  const startDuplicate = (idx) => {
    setDuplicateIdx(idx);
    setDuplicateName(projects[idx]?.name ? projects[idx].name + '_copy' : '');
  };

  const cancelDuplicate = () => {
    setDuplicateIdx(null);
    setDuplicateName('');
  };

  const duplicateProject = async (idx) => {
    const source = projects[idx]?.name;
    const newName = duplicateName.trim();
    if (!source || !newName) return;
    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects/duplicate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ source_name: source, new_name: newName }),
      });
      if (res.ok) {
        cancelDuplicate();
        fetchProjects();
      } else if (res.status === 409) {
        const text = await res.text();
        alert(text || 'Destination project already exists');
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      } else {
        const text = await res.text();
        alert(text || 'Failed to duplicate project');
      }
    } catch (e) {
      console.error('duplicate project failed', e);
      alert('Unexpected error duplicating project');
    }
  };

  const openAudit = async (name) => {
    if (!name) return;
    setAuditProject(name);
    setAuditOpen(true);
    setAuditLoading(true);
    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects/audit?project=${encodeURIComponent(name)}`);
      if (res.ok) {
        const data = await res.json();
        setAuditEntries(Array.isArray(data) ? data : []);
      } else if (res.status === 401) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      } else {
        const text = await res.text();
        alert(text || 'Failed to load project audit');
        setAuditEntries([]);
      }
    } catch (e) {
      console.error('fetch project audit failed', e);
      setAuditEntries([]);
    } finally {
      setAuditLoading(false);
    }
  };

  const closeAudit = () => {
    setAuditOpen(false);
    setAuditProject('');
    setAuditEntries([]);
    setAuditLoading(false);
  };

  const deleteProject = async (idx) => {
    const name = projects[idx]?.name;
    if (!name) return;
    if (deleteConfirm.trim() !== name) return; // require exact match
    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      const res = await authenticatedFetch(`http://${host}:8080/api/projects?name=${encodeURIComponent(name)}`, {
        method: 'DELETE',
      });
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
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost';
      await authenticatedFetch(`http://${host}:8080/api/logout`, { method: 'POST' });
    } catch {}
    clearAuth();
    navigate('/');
  };

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
            <button onClick={handleLogout} className="btn btn-outline-danger btn-sm d-flex align-items-center" title="Logout">
              <LogOut className="me-1" /> Logout
            </button>
          </div>
        </div>
      </nav>

      <main className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className={`overflow-hidden transition-all duration-300 ease-in-out ${creating ? 'max-h-40 mb-8 opacity-100' : 'max-h-0 opacity-0'}`}>
          <div className="p-6 bg-white border border-indigo-100 rounded-2xl shadow-sm">
            <form onSubmit={createProject} className="flex items-end gap-4">
              <div className="flex-1">
                <input type="text" value={newProject} onChange={(e)=>setNewProject(e.target.value)} placeholder="New Project Name" className="w-full px-4 py-2 bg-gray-50 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 transition-all" autoFocus />
                <button type="submit" className="px-6 py-2 text-black font-medium rounded-full shadow-md transition-all hover:opacity-90 border-0 focus:outline-none" style={{ backgroundColor: 'skyblue' }}>Create</button>
              </div>
            </form>
          </div>
        </div>

        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-8">
          <div className="relative flex-1 md:w-64 group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400 group-focus-within:text-indigo-500 transition-colors" />
            <input type="text" placeholder="Search projects..." value={search} onChange={(e)=>setSearch(e.target.value)} className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm" />
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
          <table className="table mb-0">
            <thead>
              <tr style={{background: 'lightgray'}}>
                <th scope="col">Project</th>
                <th scope="col" className="text-end">Actions</th>
              </tr>
            </thead>
            <tbody>
              {displayed.map((p, idx) => (
                <tr key={p.name} style={{ cursor: 'pointer' }}>
                  <td onClick={() => navigate(`/project/${encodeURIComponent(p.name)}`)}>
                    {editingIdx === idx ? (
                      <input type="text" className="form-control form-control-sm" value={editingName} onChange={(e)=>setEditingName(e.target.value)} onClick={(e)=>e.stopPropagation()} onKeyDown={(e)=>{ if (e.key==='Enter') renameProject(idx); if (e.key==='Escape') cancelRename(); }} autoFocus />
                    ) : p.name}
                  </td>
                  <td className="text-end">
                    {editingIdx === idx ? (
                      <>
                        <button className="btn btn-sm btn-success me-2" onClick={(ev)=>{ev.stopPropagation(); renameProject(idx);}}>Save</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelRename();}}>Cancel</button>
                      </>
                    ) : deleteIdx === idx ? (
                      <div className="d-inline-flex align-items-center gap-2">
                        <input type="text" className="form-control form-control-sm" placeholder={`Type '${p.name}'`} value={deleteConfirm} onChange={(e)=>setDeleteConfirm(e.target.value)} style={{ maxWidth: 200 }} />
                        <button className="btn btn-sm btn-danger" disabled={deleteConfirm.trim()!==p.name} onClick={(ev)=>{ev.stopPropagation(); deleteProject(idx);}}>Confirm</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelDelete();}}>Cancel</button>
                      </div>
                    ) : duplicateIdx === idx ? (
                      <div className="d-inline-flex align-items-center gap-2">
                        <input type="text" className="form-control form-control-sm" placeholder={`New project name`} value={duplicateName} onChange={(e)=>setDuplicateName(e.target.value)} style={{ maxWidth: 220 }} />
                        <button className="btn btn-sm btn-success" disabled={!duplicateName.trim()} onClick={(ev)=>{ev.stopPropagation(); duplicateProject(idx);}}>Duplicate</button>
                        <button className="btn btn-sm btn-secondary" onClick={(ev)=>{ev.stopPropagation(); cancelDuplicate();}}>Cancel</button>
                      </div>
                    ) : (
                      <>
                        <button className="btn btn-sm btn-outline-primary me-2" onClick={(ev)=>{ev.stopPropagation(); startRename(idx);}}>
                          <Edit2 size={14} className="me-1"/> Rename
                        </button>
                        <button className="btn btn-sm btn-outline-primary me-2" onClick={(ev)=>{ev.stopPropagation(); startDuplicate(idx);}}>
                          <FolderPlus size={14} className="me-1"/> Duplicate
                        </button>
                        <button className="btn btn-sm btn-outline-secondary me-2" onClick={(ev)=>{ev.stopPropagation(); openAudit(p.name);} }>
                          Audit
                        </button>
                        <button className="btn btn-sm btn-outline-danger" onClick={(ev)=>{ev.stopPropagation(); requestDelete(idx);}}>
                          <Trash2 size={14} className="me-1"/> Delete
                        </button>
                      </>
                    )}
                  </td>
                </tr>
              ))}
              {displayed.length === 0 && (
                <tr><td colSpan="2" className="text-center text-muted py-4">No projects found.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </main>
      {auditOpen && (
        <div className="position-fixed top-0 start-0 w-100 h-100" style={{ backgroundColor: 'rgba(0,0,0,0.4)' }} onClick={closeAudit}>
          <div className="container h-100 d-flex align-items-center justify-content-center" onClick={(e)=>e.stopPropagation()}>
            <div className="bg-white rounded-2xl shadow-lg border border-gray-200" style={{ maxWidth: '800px', width: '100%' }}>
              <div className="p-3 border-bottom d-flex justify-content-between align-items-center">
                <h5 className="mb-0">Project Audit: {auditProject}</h5>
                <button className="btn btn-sm btn-outline-secondary" onClick={closeAudit}>Close</button>
              </div>
              <div className="p-3" style={{ maxHeight: '60vh', overflowY: 'auto' }}>
                {auditLoading ? (
                  <div className="text-center text-muted py-4">Loading…</div>
                ) : auditEntries.length === 0 ? (
                  <div className="text-center text-muted py-4">No audit entries.</div>
                ) : (
                  <table className="table table-sm mb-0">
                    <thead>
                      <tr>
                        <th style={{width:'24%'}}>Time</th>
                        <th style={{width:'16%'}}>User</th>
                        <th style={{width:'20%'}}>Action</th>
                        <th>Details</th>
                      </tr>
                    </thead>
                    <tbody>
                      {auditEntries.map((e, i) => (
                        <tr key={i}>
                          <td>{formatTimestamp(e.timestamp)}</td>
                          <td>{e.user}</td>
                          <td>{e.action}</td>
                          <td>{e.details}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
