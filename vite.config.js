import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { nodePolyfills } from 'vite-plugin-node-polyfills'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    nodePolyfills({
      // .doc 解析依赖 Node 内建模块（fs/stream/path 等），浏览器端需垫片
      include: ['fs', 'stream', 'path', 'buffer', 'events', 'util', 'process', 'zlib']
    })
  ],
  base: './',
  build: {
    target: 'es2020'
  }
})
