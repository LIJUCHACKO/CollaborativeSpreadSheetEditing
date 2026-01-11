import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import './bootstrap/dist/css/bootstrap.min.css';

export default function Login() {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [isRegistering, setIsRegistering] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const navigate = useNavigate();

  const handleLogin = async (e) => {
    e.preventDefault();
    setError('');

    if (!username.trim() || !password.trim()) {
      setError("Username and password are required.");
      return;
    }

    // Removed Terms & Privacy Policy checkbox and related validation

    setLoading(true);
    const endpoint = isRegistering ? '/api/register' : '/api/login';

    try {
      const host = import.meta.env.VITE_BACKEND_HOST || 'localhost  ';
      const res = await fetch(`http://${host}:8080${endpoint}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
      });

      if (res.ok) {
        if (isRegistering) {
          setIsRegistering(false);
          setError("Registration successful! Please log in.");
          setUsername('');
          setPassword('');
        } else {
          const data = await res.json();
          // Store token and login timestamp
          localStorage.setItem('auth_token', data.token);
          localStorage.setItem('chat_username', data.username);
          localStorage.setItem('login_time', new Date().getTime().toString());
          navigate('/dashboard');
        }
      } else {
        const text = await res.text();
        setError(text || "Authentication failed. Please try again.");
      }
    } catch (e) {
      setError("Network error. Please check your connection and try again.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="text-center">
      <form className="form-signin" onSubmit={handleLogin}>
        <h2 className="h3 mb-3 font-weight-normal text-center">
          {isRegistering ? 'Register' : 'Login'}
        </h2>

        {error && (
          <div className="alert alert-danger text-center" role="alert" style={{ maxWidth: '400px', margin: '0 auto' }}>
            {error}
          </div>
        )}

        <div className="mb-3" style={{ maxWidth: '400px', margin: '0 auto' }}>
          <label htmlFor="username" className="form-label">Username</label>
          <div className="input-group">
            {/* Username Icon */}
            <span className="input-group-text">
              <i
                className="bi bi-person"
                aria-hidden="true"
                style={{ backgroundColor: '#000', color: '#fff', borderRadius: '4px', padding: '4px' }}
              ></i>
            </span>
            <input
              type="text"
              id="username"
              className="form-control"
              placeholder="Enter your username"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              required
              aria-label="Username"
            />
          </div>
        </div>

        <div className="mb-3" style={{ maxWidth: '400px', margin: '0 auto' }}>
          <label htmlFor="password" className="form-label">Password</label>
          <div className="input-group">
            <span className="input-group-text">
              <i
                className="bi bi-lock"
                aria-hidden="true"
                style={{ backgroundColor: '#000', color: '#fff', borderRadius: '4px', padding: '4px' }}
              ></i>
            </span>
            <input
              type={showPassword ? "text" : "password"}
              id="password"
              className="form-control"
              placeholder="Enter your password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              aria-label="Password"
            />
            <button
              type="button"
              className="btn btn-outline-secondary"
              onClick={() => setShowPassword(!showPassword)}
              aria-label={showPassword ? "Hide password" : "Show password"}
            >
              <i
                className={`bi ${showPassword ? 'bi-eye-slash' : 'bi-eye'}`}
                style={{ backgroundColor: '#000', color: '#fff', borderRadius: '4px', padding: '4px' }}
              ></i>
            </button>
          </div>
        </div>

        {/* Terms & Privacy Policy checkbox removed */}

        

        <button
          type="submit"
          className="btn btn-primary w-100"
          disabled={loading}
          style={{ maxWidth: '400px', margin: '0 auto' }}
        >
          {loading ? (
            <span className="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>
          ) : (
            isRegistering ? 'Create Account' : 'Sign In'
          )}
        </button>
          {!isRegistering && (
          <div className="text-end mb-3">
            <button
              type="button"
              className="btn btn-link text-decoration-none"
              onClick={() => setIsRegistering(true)}
            >
              Register
            </button>
          </div>
        )}
        {isRegistering && (
          <div className="text-center mt-3" style={{ maxWidth: '400px', margin: '0 auto' }}>
            <button
              type="button"
              className="btn btn-link text-decoration-none"
              onClick={() => setIsRegistering(false)}
            >
              Already have an account? Login
            </button>
          </div>
        )}
      </form>
    </main>
  );
}
