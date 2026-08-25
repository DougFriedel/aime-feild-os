import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],

  // The repo's static folder is "Public" with a capital P. Vite defaults to
  // lowercase "public", and Netlify builds on Linux where that's a different
  // folder — so nothing in it was reaching the deployed site. Pointing publicDir
  // at the real folder name fixes that without a case-only rename in git, which
  // git handles badly.
  publicDir: 'Public',

  build: {
    rollupOptions: {
      // Externalize Capacitor packages — they're only available inside
      // the native app shell, not needed for the web/Netlify build
      external: [
        '@capacitor/core',
        '@capacitor/cli',
        '@capacitor/ios',
        '@capacitor/android',
        '@capacitor/push-notifications',
        '@capacitor/network',
        '@capacitor/splash-screen',
        '@capacitor/status-bar',
        '@capacitor/local-notifications',
        '@capacitor/assets',
      ],
    },
  },
})
