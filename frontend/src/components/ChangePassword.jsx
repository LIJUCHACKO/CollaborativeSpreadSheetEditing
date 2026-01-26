import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { authenticatedFetch, isSessionValid, clearAuth, getUsername, apiUrl } from '../utils/auth';
import { ArrowLeft, Lock, Save, User } from 'lucide-react';
import 'bootstrap/dist/css/bootstrap.min.css';

export default function ChangePassword() {
  const navigate = useNavigate();
  const username = getUsername();

  const [oldPwd, setOldPwd] = useState('');
  const [newPwd, setNewPwd] = useState('');
  const [confirmPwd, setConfirmPwd] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!username || !isSessionValid()) {
      clearAuth();
      navigate('/');
      return;
    }
  }, [username, navigate]);

  const submit = async (e) => {
    e.preventDefault();
    if (!oldPwd || !newPwd || !confirmPwd) {
      alert('Please fill all password fields.');
      return;
    }
    if (newPwd !== confirmPwd) {
      alert('New password and confirm password do not match.');
      return;
    }
    if (newPwd.length < 6) {
      alert('New password must be at least 6 characters.');
      return;
    }
    try {
      setBusy(true);
      const res = await authenticatedFetch(apiUrl('/api/user/password'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ old_password: oldPwd, new_password: newPwd })
      });
      if (!res.ok) {
        const text = await res.text();
        alert(text || 'Failed to change password');
        return;
      }
      alert('Password updated successfully');
      navigate('/projects');
    } catch (err) {
      console.error('Password change error', err);
      alert('Unexpected error changing password');
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
      <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
        <div className="container-fluid">
          <button onClick={() => navigate('/projects')} className="btn btn-outline-primary btn-sm d-flex align-items-center">
            <ArrowLeft className="me-1" />
          </button>
          <span className="navbar-text ms-2 d-flex align-items-center fw-bold">
            <Lock className="me-2" /> Change Password
          </span>
          <div className="ms-auto d-flex align-items-center">
            <span className="navbar-text me-3 d-flex align-items-center">
              <User className="me-1" /> {username}
            </span>
          </div>
        </div>
      </nav>

      <main className="flex-1 max-w-xl w-full mx-auto px-4 py-8">
        <div className="bg-white border border-gray-200 rounded-2xl shadow-sm p-4">
          <form onSubmit={submit}>
            <div className="mb-3">
              <label className="form-label">Current Password</label>
              <input type="password" className="form-control" value={oldPwd} onChange={(e)=>setOldPwd(e.target.value)} autoFocus />
            </div>
            <div className="mb-3">
              <label className="form-label">New Password</label>
              <input type="password" className="form-control" value={newPwd} onChange={(e)=>setNewPwd(e.target.value)} />
            </div>
            <div className="mb-3">
              <label className="form-label">Confirm New Password</label>
              <input type="password" className="form-control" value={confirmPwd} onChange={(e)=>setConfirmPwd(e.target.value)} />
            </div>
            <div className="d-flex justify-content-end gap-2">
              <button type="button" className="btn btn-sm btn-secondary" onClick={() => navigate('/projects')}>Cancel</button>
              <button type="submit" className="btn btn-sm btn-outline-primary d-flex align-items-center" disabled={busy}>
                <Save size={14} className="me-1"/> Update Password
              </button>
            </div>
            <p className="text-muted mt-2">Password must be at least 6 characters.</p>
          </form>
        </div>
      </main>
    </div>
  );
}
