import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
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
