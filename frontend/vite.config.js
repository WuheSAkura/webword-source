import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  server: {
    port: 3000,
    proxy: {
      '/api': {
        // Windows 上 localhost 可能优先解析到 ::1，而本地 FastAPI 通常只监听 IPv4。
        // 本地开发默认 8010：8001 常被其他项目（如 xianyu_backend Docker）占用。
        target: process.env.VITE_API_PROXY_TARGET || 'http://127.0.0.1:8010',
        changeOrigin: true,
      },
    },
  },
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
  },
})
