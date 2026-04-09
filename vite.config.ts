import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig(({ command }) => ({
  plugins: [react()],
  // Use relative asset paths for production so the app works on GitHub Pages
  // regardless of whether it is served from a repository subpath or a custom domain.
  base: command === 'build' ? './' : '/',
}))
