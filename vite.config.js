import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // In local dev, forwards to `npm run dev:api` (see local-api-server.mjs)
      // so the frontend can call relative /api/... paths exactly like it
      // does once deployed on Vercel - no VITE_API_URL needed either way.
      '/api': 'http://localhost:4000',
    },
  },
})
