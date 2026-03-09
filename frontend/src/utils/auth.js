// Authentication utility functions

const SESSION_TIMEOUT = 60 * 60 * 1000; // 1 hour in milliseconds

/**
 * Check if the current session is still valid (within 1 hour timeout)
 * @returns {boolean} true if session is valid, false otherwise
 */
export function isSessionValid() {
  const token = localStorage.getItem('auth_token');
  const loginTime = localStorage.getItem('login_time');
  
  if (!token || !loginTime) {
    return false;
  }
  
  const currentTime = new Date().getTime();
  const elapsedTime = currentTime - parseInt(loginTime);
  
  return elapsedTime < SESSION_TIMEOUT;
}

/**
 * Get the authentication token
 * @returns {string|null} token or null if not found
 */
export function getAuthToken() {
  return localStorage.getItem('auth_token');
}

/**
 * Get the current username
 * @returns {string|null} username or null if not found
 */
export function getUsername() {
  return localStorage.getItem('chat_username');
}

/**
 * Clear all authentication data from localStorage
 */
export function clearAuth() {
  localStorage.removeItem('auth_token');
  localStorage.removeItem('chat_username');
  localStorage.removeItem('login_time');
  localStorage.removeItem('is_admin');
  localStorage.removeItem('can_create_project');
}

/**
 * Get remaining session time in milliseconds
 * @returns {number} remaining time in milliseconds, or 0 if expired
 */
export function getRemainingSessionTime() {
  const loginTime = localStorage.getItem('login_time');
  
  if (!loginTime) {
    return 0;
  }
  
  const currentTime = new Date().getTime();
  const elapsedTime = currentTime - parseInt(loginTime);
  const remaining = SESSION_TIMEOUT - elapsedTime;
  
  return remaining > 0 ? remaining : 0;
}

/**
 * Make an authenticated API request
 * @param {string} url - The API endpoint URL
 * @param {object} options - Fetch options
 * @returns {Promise<Response>} The fetch response
 */
export async function authenticatedFetch(url, options = {}) {
  const token = getAuthToken();
  
  if (!token) {
    throw new Error('No authentication token found');
  }
  
  const headers = {
    ...options.headers,
    'Authorization': token,
  };
  
  return fetch(url, {
    ...options,
    headers,
  });
}

/**
 * Build full backend URL from env, accepting both host or full URL.
 * When VITE_BACKEND_HOST is NOT set, returns a relative path (e.g. "/api/login")
 * so the Vite dev-server proxy forwards the request to the backend, avoiding CORS.
 * Examples:
 * - VITE_BACKEND_HOST not set            => "/api/login"  (relative, goes through proxy → localhost:8082)
 * - VITE_BACKEND_HOST="192.168.0.102"    => "http://192.168.0.102:8082/api/login"
 * - VITE_BACKEND_HOST="127.0.0.1:9090"  => "http://127.0.0.1:9090/api/login"
 * - VITE_BACKEND_HOST="http://host:8082" => "http://host:8082/api/login"
 * @param {string} path like "/api/login"
 * @returns {string} full or relative URL
 */
export function apiUrl(path) {
  const raw = import.meta?.env?.VITE_BACKEND_HOST;
  const p = path.startsWith('/') ? path : `/${path}`;
  // No env var set → use relative URL so Vite proxy handles the request
  if (!raw || raw.trim() === '') {
    return p;
  }
  let base;
  if (raw.startsWith('http://') || raw.startsWith('https://')) {
    base = raw.replace(/\/$/, '');
  } else {
    // treat as host[:port]
    base = raw.includes(':') ? `http://${raw}` : `http://${raw}:8082`;
  }
  return `${base}${p}`;
}

/**
 * Returns true if the current logged-in user is an admin.
 */
export function isAdmin() {
  return localStorage.getItem('is_admin') === 'true';
}

/**
 * Returns true if the current user is allowed to create projects.
 */
export function canCreateProject() {
  return localStorage.getItem('can_create_project') === 'true';
}
