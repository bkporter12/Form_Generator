import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3002, // You can change this to any 4-digit number you want
    strictPort: true, // This tells Vite to crash if 3000 is also taken, rather than silently guessing a new one
  }
})