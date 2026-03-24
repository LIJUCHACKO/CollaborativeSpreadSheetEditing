import React, { useEffect, useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { authenticatedFetch, isSessionValid, clearAuth, getUsername, isAdmin, apiUrl } from '../utils/auth';
import { ShieldCheck, KeyRound, ToggleLeft, ToggleRight, LogOut, ArrowLeft, RefreshCw, Download, AlertTriangle, CheckCircle } from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';

export default function Admin() {
  const navigate = useNavigate();
  const username = getUsername();

  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Password reset state
  const [pwTarget, setPwTarget] = useState(null);   // username being reset
  const [newPw, setNewPw] = useState('');
  const [pwMsg, setPwMsg] = useState('');

  // Ownership transfer state
  const [projectName, setProjectName] = useState('');
  const [projectNewOwner, setProjectNewOwner] = useState('');
  const [projectTransferMsg, setProjectTransferMsg] = useState('');

  const [sheetProject, setSheetProject] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [sheetNewOwner, setSheetNewOwner] = useState('');
  const [sheetTransferMsg, setSheetTransferMsg] = useState('');

  // Integrity report state
  const [integrityReport, setIntegrityReport] = useState(null);
  const [integrityLoading, setIntegrityLoading] = useState(false);

  // LLM settings state
  const [llmUrl, setLlmUrl] = useState('');
  const [llmUrlSaved, setLlmUrlSaved] = useState('');
  const [llmMsg, setLlmMsg] = useState('');

  useEffect(() => {
    if (!username || !isSessionValid()) {
      clearAuth();
      navigate('/');
      return;
    }
    if (!isAdmin()) {
      navigate('/projects');
      return;
    }
    fetchUsers();
    fetchIntegrityReport();
    fetchLLMSettings();

    const interval = setInterval(() => {
      if (!isSessionValid()) {
        clearAuth();
        alert('Your session has expired. Please log in again.');
        navigate('/');
      }
    }, 60000);
    return () => clearInterval(interval);
  }, [username, navigate]);

  const fetchUsers = async () => {
    setLoading(true);
    setError('');
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/users'));
      if (res.ok) {
        const data = await res.json();
        // Sort: admin first, then alphabetically
        data.sort((a, b) => {
          if (a.is_admin && !b.is_admin) return -1;
          if (!a.is_admin && b.is_admin) return 1;
          return a.username.localeCompare(b.username);
        });
        setUsers(data);
      } else if (res.status === 401) {
        clearAuth();
        navigate('/');
      } else if (res.status === 403) {
        navigate('/projects');
      } else {
        setError('Failed to load users');
      }
    } catch (e) {
      setError('Network error');
    } finally {
      setLoading(false);
    }
  };

  const fetchIntegrityReport = async () => {
    setIntegrityLoading(true);
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/integrity'));
      if (res.ok) {
        const data = await res.json();
        // Sort: corrupted first, then by path
        if (data.files) {
          data.files.sort((a, b) => {
            if (!a.intact && !a.missing && (b.intact || b.missing)) return -1;
            if (!b.intact && !b.missing && (a.intact || a.missing)) return 1;
            return (a.path || '').localeCompare(b.path || '');
          });
        }
        setIntegrityReport(data);
      }
    } catch (e) {
      console.error('integrity fetch failed', e);
    } finally {
      setIntegrityLoading(false);
    }
  };

  const fetchLLMSettings = async () => {
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/llm'));
      if (res.ok) {
        const data = await res.json();
        setLlmUrl(data.url || '');
        setLlmUrlSaved(data.url || '');
      }
    } catch (e) {
      console.error('LLM settings fetch failed', e);
    }
  };

  const saveLLMSettings = async () => {
    setLlmMsg('');
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/llm'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: llmUrl.trim() }),
      });
      if (res.ok) {
        setLlmUrlSaved(llmUrl.trim());
        setLlmMsg('LLM URL saved successfully');
        setTimeout(() => setLlmMsg(''), 3000);
      } else {
        const text = await res.text();
        setLlmMsg(text || 'Failed to save LLM URL');
      }
    } catch (e) {
      setLlmMsg('Network error');
    }
  };

  const togglePermission = async (user) => {
    if (user.is_admin) return; // cannot change admin's own permission
    const newVal = !user.can_create_project;
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/user/permission'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: user.username, can_create_project: newVal }),
      });
      if (res.ok) {
        setUsers(prev =>
          prev.map(u => u.username === user.username ? { ...u, can_create_project: newVal } : u)
        );
      } else {
        const text = await res.text();
        alert(text || 'Failed to update permission');
      }
    } catch (e) {
      alert('Network error');
    }
  };

  const startResetPassword = (uname) => {
    setPwTarget(uname);
    setNewPw('');
    setPwMsg('');
  };

  const cancelResetPassword = () => {
    setPwTarget(null);
    setNewPw('');
    setPwMsg('');
  };

  const confirmResetPassword = async () => {
    if (!pwTarget || !newPw.trim()) return;
    if (newPw.trim().length < 6) {
      setPwMsg('Password must be at least 6 characters');
      return;
    }
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/user/password'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: pwTarget, new_password: newPw.trim() }),
      });
      if (res.ok) {
        setPwMsg('Password updated successfully');
        setTimeout(() => cancelResetPassword(), 1500);
      } else {
        const text = await res.text();
        setPwMsg(text || 'Failed to update password');
      }
    } catch (e) {
      setPwMsg('Network error');
    }
  };

  const downloadBackup = async () => {
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/backup'));
      if (!res.ok) {
        const text = await res.text();
        alert(text || 'Failed to download backup');
        return;
      }
      const blob = await res.blob();
      // Extract filename from Content-Disposition header if available
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = disposition.match(/filename="?([^"]+)"?/);
      const filename = match ? match[1] : 'backup.zip';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert('Network error while downloading backup');
    }
  };

  const handleLogout = async () => {
    try {
      await authenticatedFetch(apiUrl('/api/logout'), { method: 'POST' });
    } catch (_) {}
    clearAuth();
    navigate('/');
  };

  const submitProjectTransfer = async (e) => {
    e.preventDefault();
    setProjectTransferMsg('');
    const proj = projectName.trim();
    const owner = projectNewOwner.trim();
    if (!proj || !owner) {
      setProjectTransferMsg('Project and new owner are required');
      return;
    }
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/project/transfer'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ project: proj, new_owner: owner }),
      });
      if (res.ok) {
        setProjectTransferMsg('Project owner updated successfully');
      } else {
        const text = await res.text();
        setProjectTransferMsg(text || 'Failed to update project owner');
      }
    } catch (err) {
      setProjectTransferMsg('Network error');
    }
  };

  const submitSheetTransfer = async (e) => {
    e.preventDefault();
    setSheetTransferMsg('');
    const proj = sheetProject.trim();
    const sname = sheetName.trim();
    const owner = sheetNewOwner.trim();
    if (!sname || !owner) {
      setSheetTransferMsg('Sheet name and new owner are required');
      return;
    }
    try {
      const res = await authenticatedFetch(apiUrl('/api/admin/sheet/transfer'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ project: proj, sheet_name: sname, new_owner: owner }),
      });
      if (res.ok) {
        setSheetTransferMsg('Sheet owner updated successfully');
      } else {
        const text = await res.text();
        setSheetTransferMsg(text || 'Failed to update sheet owner');
      }
    } catch (err) {
      setSheetTransferMsg('Network error');
    }
  };

  return (
    <div className="min-vh-100 bg-light">
      {/* Navbar */}
      <nav className="navbar navbar-dark bg-dark px-4 py-2 d-flex justify-content-between align-items-center">
        <div className="d-flex align-items-center gap-3">
          <Link to="/projects" className="btn btn-sm btn-outline-light">
            <ArrowLeft size={14} className="me-1" /> Projects
          </Link>
          <span className="navbar-brand mb-0 d-flex align-items-center gap-2">
            <ShieldCheck size={20} /> Admin Panel
          </span>
        </div>
        <div className="d-flex align-items-center gap-2">
          <span className="text-light small me-2">{username}</span>
          <button className="btn btn-sm btn-outline-info" onClick={downloadBackup} title="Download full DATA folder as zip backup">
            <Download size={14} className="me-1" /> Backup
          </button>
          <button className="btn btn-sm btn-outline-light" onClick={handleLogout}>
            <LogOut size={14} className="me-1" /> Logout
          </button>
        </div>
      </nav>

      <div className="container py-4" style={{ maxWidth: 900 }}>
        <div className="d-flex align-items-center justify-content-between mb-4">
          <h4 className="mb-0 fw-bold">User Management</h4>
          <button className="btn btn-sm btn-outline-secondary" onClick={fetchUsers}>
            <RefreshCw size={14} className="me-1" /> Refresh
          </button>
        </div>

        {error && <div className="alert alert-danger">{error}</div>}

        {loading ? (
          <div className="text-center py-5">
            <div className="spinner-border text-secondary" role="status" />
          </div>
        ) : (
          <div className="card shadow-sm border-0">
            <table className="table table-hover mb-0 align-middle">
              <thead className="table-dark">
                <tr>
                  <th>Username</th>
                  <th>Role</th>
                  <th className="text-center">Can Create Project</th>
                  <th className="text-end">Actions</th>
                </tr>
              </thead>
              <tbody>
                {users.map(user => (
                  <React.Fragment key={user.username}>
                    <tr>
                      <td>
                        <span className="fw-semibold">{user.username}</span>
                        {user.username === username && (
                          <span className="badge bg-info ms-2 text-dark" style={{ fontSize: '0.7rem' }}>you</span>
                        )}
                      </td>
                      <td>
                        {user.is_admin ? (
                          <span className="badge bg-danger">Admin</span>
                        ) : (
                          <span className="badge bg-secondary">User</span>
                        )}
                      </td>
                      <td className="text-center">
                        {user.is_admin ? (
                          <span className="text-muted small">always</span>
                        ) : (
                          <button
                            className={`btn btn-sm ${user.can_create_project ? 'btn-success' : 'btn-outline-secondary'}`}
                            onClick={() => togglePermission(user)}
                            title={user.can_create_project ? 'Revoke permission' : 'Grant permission'}
                          >
                            {user.can_create_project ? (
                              <><ToggleRight size={16} className="me-1" /> Allowed</>
                            ) : (
                              <><ToggleLeft size={16} className="me-1" /> Blocked</>
                            )}
                          </button>
                        )}
                      </td>
                      <td className="text-end">
                        {!user.is_admin && (
                          <button
                            className="btn btn-sm btn-outline-warning"
                            onClick={() => startResetPassword(user.username)}
                          >
                            <KeyRound size={14} className="me-1" /> Reset Password
                          </button>
                        )}
                      </td>
                    </tr>

                    {/* Inline password reset form */}
                    {pwTarget === user.username && (
                      <tr className="table-warning">
                        <td colSpan={4}>
                          <div className="d-flex align-items-center gap-2 flex-wrap">
                            <span className="small text-muted">New password for <strong>{pwTarget}</strong>:</span>
                            <input
                              type="password"
                              className="form-control form-control-sm"
                              style={{ maxWidth: 260 }}
                              placeholder="New password (min 6 chars)"
                              value={newPw}
                              onChange={e => { setNewPw(e.target.value); setPwMsg(''); }}
                              onKeyDown={e => { if (e.key === 'Enter') confirmResetPassword(); if (e.key === 'Escape') cancelResetPassword(); }}
                              autoFocus
                            />
                            <button className="btn btn-sm btn-success" disabled={!newPw.trim()} onClick={confirmResetPassword}>
                              Set Password
                            </button>
                            <button className="btn btn-sm btn-secondary" onClick={cancelResetPassword}>
                              Cancel
                            </button>
                            {pwMsg && (
                              <span className={`small ms-1 ${pwMsg.includes('success') ? 'text-success' : 'text-danger'}`}>
                                {pwMsg}
                              </span>
                            )}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                ))}

                {users.length === 0 && (
                  <tr>
                    <td colSpan={4} className="text-center text-muted py-4">No users found.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Ownership transfer section */}
        <div className="row mt-4 g-3">
          <div className="mb-0 mt-1">
            <div className="card shadow-sm border-0 h-100">
              <div className="card-body">
                <h5 className="card-title">Transfer Project Ownership</h5>
                <p className="text-muted small mb-3">Change the owner of an entire project.</p>
                <form onSubmit={submitProjectTransfer} className="vstack gap-2">
                  <input
                    type="text"
                    className="form-control form-control-sm"
                    placeholder="Project name (folder name in DATA)"
                    value={projectName}
                    onChange={e => setProjectName(e.target.value)}
                  />
                  <input
                    type="text"
                    className="form-control form-control-sm"
                    placeholder="New owner username"
                    value={projectNewOwner}
                    onChange={e => setProjectNewOwner(e.target.value)}
                  />
                  <button type="submit" className="btn btn-sm btn-primary align-self-start">
                    Transfer Project Owner
                  </button>
                  {projectTransferMsg && (
                    <div className="small mt-1 {projectTransferMsg.includes('successfully') ? 'text-success' : 'text-danger'}">
                      {projectTransferMsg}
                    </div>
                  )}
                </form>
              </div>
            </div>
          </div>

          
        </div>

        {/* Integrity Report */}
        <div className="mt-4 card shadow-sm border-0">
          <div className="card-header d-flex align-items-center justify-content-between bg-white">
            <h5 className="mb-0 fw-bold d-flex align-items-center gap-2">
              {integrityReport?.system_corrupt
                ? <AlertTriangle size={18} className="text-danger" />
                : <CheckCircle size={18} className="text-success" />}
              File Integrity Report
            </h5>
            <button className="btn btn-sm btn-outline-secondary" onClick={fetchIntegrityReport} disabled={integrityLoading}>
              <RefreshCw size={14} className="me-1" /> Refresh
            </button>
          </div>
          <div className="card-body p-0">
            {integrityLoading ? (
              <div className="text-center py-3"><div className="spinner-border spinner-border-sm text-secondary" /></div>
            ) : integrityReport ? (
              <>
                {integrityReport.system_corrupt && (
                  <div className="alert alert-danger mb-0 rounded-0 d-flex align-items-center gap-2 px-3 py-2">
                    <AlertTriangle size={16} />
                    <span><strong>Critical:</strong> users.json or projects.json is corrupt — all projects are read-only.</span>
                  </div>
                )}
                <table className="table table-sm table-hover mb-0 align-middle">
                  <thead className="table-light">
                    <tr>
                      <th>File</th>
                      <th className="text-center">Status</th>
                      <th>Details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(integrityReport.files || []).map(f => (
                      <tr key={f.path}>
                        <td className="font-monospace small">{f.path}</td>
                        <td className="text-center">
                          {f.intact
                            ? <span className="badge bg-success d-inline-flex align-items-center gap-1"><CheckCircle size={11}/> Intact</span>
                            : f.missing
                              ? <span className="badge bg-warning text-dark d-inline-flex align-items-center gap-1"><AlertTriangle size={11}/> No checksum</span>
                              : <span className="badge bg-danger d-inline-flex align-items-center gap-1"><AlertTriangle size={11}/> Corrupt</span>
                          }
                        </td>
                        <td className="small text-muted">{f.error || (f.missing ? 'First run or external write — will be protected on next save' : '')}</td>
                      </tr>
                    ))}
                    {(!integrityReport.files || integrityReport.files.length === 0) && (
                      <tr><td colSpan={3} className="text-center text-muted py-3">No files tracked yet.</td></tr>
                    )}
                  </tbody>
                </table>
              </>
            ) : (
              <div className="text-center text-muted py-3 small">Click Refresh to load integrity report.</div>
            )}
          </div>
        </div>

        {/* LLM Settings */}
        <div className="card mt-4 shadow-sm">
          <div className="card-header bg-white d-flex align-items-center gap-2 py-2 px-3">
            <strong className="small">LLM Server Settings</strong>
          </div>
          <div className="card-body px-3 py-3">
            <div className="mb-2 small text-muted">
              Configure the OpenAI-compatible LLM server URL used by AI Generated cells. Leave empty to disable AI cell execution.
            </div>
            <div className="d-flex gap-2 align-items-center">
              <input
                type="text"
                className="form-control form-control-sm"
                placeholder="e.g. http://localhost:8080"
                value={llmUrl}
                onChange={e => setLlmUrl(e.target.value)}
                style={{ maxWidth: 400 }}
              />
              <button
                className="btn btn-sm btn-primary"
                onClick={saveLLMSettings}
                disabled={llmUrl.trim() === llmUrlSaved}
              >
                Save
              </button>
            </div>
            {llmMsg && <div className="mt-2 small text-success">{llmMsg}</div>}
          </div>
        </div>

        <div className="mt-4 p-3 bg-white border rounded small text-muted">
          <strong>Notes:</strong>
          <ul className="mb-0 mt-1">
            <li>New users cannot create projects until an admin grants them permission.</li>
            <li>Admin password can be changed via the <Link to="/change-password">Change Password</Link> page.</li>
            <li>The admin account cannot be deleted or demoted.</li>
          </ul>
        </div>
      </div>
    </div>
  );
}
