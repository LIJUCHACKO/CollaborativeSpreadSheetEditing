import React, { useEffect, useState } from 'react';
import { useNavigate, useParams, useLocation } from 'react-router-dom';
import { isSessionValid, clearAuth, getUsername, authenticatedFetch, apiUrl } from '../utils/auth';
import { ArrowLeft, Settings as SettingsIcon, User, Save, Lock } from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';


export default function Settings() {
  const { id } = useParams();
  const navigate = useNavigate();
  const location = useLocation();
  const username = getUsername();
  const project = (() => {
    try {
      const params = new URLSearchParams(location.search);
      return params.get('project') || '';
    } catch {
      return '';
    }
  })();

  const [sheet, setSheet] = useState(null);
  const [users, setUsers] = useState([]);
  const [editors, setEditors] = useState([]);
  const [newOwner, setNewOwner] = useState('');
  const isOwner = sheet && sheet.owner === username;

  useEffect(() => {
    if (!username || !isSessionValid()) {
      clearAuth();
      navigate('/');
      return;
    }

    const fetchData = async () => {
      try {
        const sheetRes = await authenticatedFetch(apiUrl(`/api/sheet?id=${encodeURIComponent(id)}${project ? `&project=${encodeURIComponent(project)}` : ''}`));
        if (!sheetRes.ok) {
          if (sheetRes.status === 401) {
            clearAuth();
            alert('Your session has expired. Please log in again.');
            navigate('/');
            return;
          }
          const text = await sheetRes.text();
          alert(text || 'Failed to fetch sheet');
          return;
        }
        const s = await sheetRes.json();
        setSheet(s);
        setEditors((s.permissions?.editors) || []);
        setNewOwner(s.owner || '');

        const usersRes = await authenticatedFetch(apiUrl('/api/users'));
        if (usersRes.ok) {
          const list = await usersRes.json();
          setUsers(Array.isArray(list) ? list : []);
        }
      } catch (e) {
        console.error('Settings fetch error', e);
      }
    };
    fetchData();
  }, [id, username, navigate]);

  const toggleItem = (list, item) => {
    const exists = list.includes(item);
    return exists ? list.filter(x => x !== item) : [...list, item];
  };

  const savePermissions = async () => {
    if (!isOwner) return;
    try {
      const res = await authenticatedFetch(apiUrl(`/api/sheet/permissions?sheet_id=${encodeURIComponent(id)}${project ? `&project=${encodeURIComponent(project)}` : ''}`), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ editors }),
      });
      if (!res.ok) {
        const text = await res.text();
        alert(text || 'Failed to update permissions');
        return;
      }
      alert('Permissions updated');
    } catch (e) {
      console.error('Save permissions error', e);
      alert('Unexpected error saving permissions');
    }
  };

  const transferOwnership = async () => {
    if (!isOwner) return;
    if (!newOwner || newOwner === sheet.owner) {
      alert('Select a different user as new owner');
      return;
    }
    try {
      const res = await authenticatedFetch(apiUrl('/api/sheet/transfer_owner'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(project ? { sheet_id: id, new_owner: newOwner, project_name: project } : { sheet_id: id, new_owner: newOwner }),
      });
      if (!res.ok) {
        const text = await res.text();
        alert(text || 'Failed to transfer ownership');
        return;
      }
      alert('Ownership transferred');
      // Reflect new owner locally
      setSheet(prev => ({ ...prev, owner: newOwner, permissions: { ...prev.permissions, editors: Array.from(new Set([...(prev.permissions?.editors || []), newOwner])) } }));
    } catch (e) {
      console.error('Transfer ownership error', e);
      alert('Unexpected error transferring ownership');
    }
  };

  if (!sheet) {
    return (
      <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
        <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
          <div className="container-fluid">
            <button onClick={() => navigate(-1)} className="btn btn-outline-primary btn-sm d-flex align-items-center">
              <ArrowLeft className="me-1" />
            </button>
            <span className="navbar-text ms-2 d-flex align-items-center fw-bold">
              <SettingsIcon className="me-2" /> Settings
            </span>
            <div className="ms-auto d-flex align-items-center">
              <span className="navbar-text me-3 d-flex align-items-center">
                <User className="me-1" /> {username}
              </span>
            </div>
          </div>
        </nav>
        <main className="flex-1 max-w-4xl w-full mx-auto px-4 py-8">Loading...</main>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
      <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
        <div className="container-fluid">
          <button onClick={() => navigate(project ? `/sheet/${id}?project=${encodeURIComponent(project)}` : `/sheet/${id}`)} className="btn btn-outline-primary btn-sm d-flex align-items-center">
            <ArrowLeft className="me-1" />
          </button>
          <span className="navbar-text ms-2 d-flex align-items-center fw-bold">
            <SettingsIcon className="me-2" /> Settings · {sheet.name || id}
          </span>
          <div className="ms-auto d-flex align-items-center">
            <span className="navbar-text me-3 d-flex align-items-center">
              <User className="me-1" /> {username}
            </span>
          </div>
        </div>
      </nav>

      <main className="flex-1 max-w-4xl w-full mx-auto px-4 py-8">
        <div className="bg-white border border-gray-200 rounded-2xl shadow-sm p-4 mb-6">
          <h5 className="mb-3">Ownership</h5>
          <p className="mb-2"><strong>Current Owner:</strong> {sheet.owner}</p>
          <div className="d-flex gap-2 align-items-center">
            <select
              className="form-select form-select-sm"
              value={newOwner}
              onChange={(e) => setNewOwner(e.target.value)}
              disabled={!isOwner}
            >
              <option value="">Select new owner</option>
              {users.map(u => (
                <option key={u} value={u}>{u}</option>
              ))}
            </select>
            <button className="btn btn-sm btn-outline-primary d-flex align-items-center" onClick={transferOwnership} disabled={!isOwner}>
              <Save size={14} className="me-1" /> Transfer
            </button>
          </div>
          {!isOwner && (
            <p className="text-muted mt-2">Only the owner can transfer ownership.</p>
          )}
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl shadow-sm p-4">
          <h5 className="mb-3">Permissions</h5>
          <div className="row">
            <div className="col-md-12 mb-3">
              <label className="form-label">Editors</label>
              <div className="d-flex flex-wrap gap-2">
                {users.map(u => (
                  <button
                    key={`editor-${u}`}
                    type="button"
                    className={`btn btn-sm ${editors.includes(u) ? 'btn-success' : 'btn-outline-secondary'}`}
                    onClick={() => isOwner && setEditors(prev => toggleItem(prev, u))}
                    disabled={!isOwner}
                  >
                    {u}
                  </button>
                ))}
              </div>
            </div>
          </div>
          <div className="d-flex justify-content-end">
            <button className="btn btn-sm btn-outline-primary d-flex align-items-center" onClick={savePermissions} disabled={!isOwner}>
              <Save size={14} className="me-1" /> Save Permissions
            </button>
          </div>
          {!isOwner && (
            <p className="text-muted mt-2">Only the owner can modify permissions.</p>
          )}
        </div>
      </main>
    </div>
  );
}
