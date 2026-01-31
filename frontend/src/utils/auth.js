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
 * Examples:
 * - VITE_BACKEND_HOST="localhost" => http://localhost:8082
 * - VITE_BACKEND_HOST="127.0.0.1:9090" => http://127.0.0.1:9090
 * - VITE_BACKEND_HOST="http://localhost:8082" => http://localhost:8082
 * - VITE_BACKEND_HOST not set => http://localhost:8082
 * @param {string} path like "/api/login"
 * @returns {string} full URL
 */
export function apiUrl(path) {
  const raw = import.meta?.env?.VITE_BACKEND_HOST;
  let base;
  if (!raw || raw.trim() === '') {
    base = 'http://localhost:8082';
  } else if (raw.startsWith('http://') || raw.startsWith('https://')) {
    base = raw.replace(/\/$/, '');
  } else {
    // treat as host[:port]
    base = raw.includes(':') ? `http://${raw}` : `http://${raw}:8082`;
  }
  const p = path.startsWith('/') ? path : `/${path}`;
  return `${base}${p}`;
}
