import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
// VITE_BACKEND_HOST can be set to point to a remote backend (e.g. "192.168.0.102:8082").
// When unset, the dev server proxies /api and /ws to localhost:8082, eliminating CORS.
const backendHost = process.env.VITE_BACKEND_HOST || 'localhost:8082';
const backendHttp = backendHost.startsWith('http') ? backendHost : `http://${backendHost}`;
const backendWs  = backendHttp.replace(/^http/, 'ws');

export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',  // accept connections from LAN, not just localhost
    port: 5175,
    proxy: {
      '/api': {
        target: backendHttp,
        changeOrigin: true,
      },
      '/ws': {
        target: backendWs,
        ws: true,
        changeOrigin: true,
      },
    },
  },
})
