import { defineConfig } from 'vite';
import { readFileSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

// Chứng chỉ https cho localhost (Word add-in khi phát triển): npm run addin:certs
function devHttps() {
  const dir = join(homedir(), '.office-addin-dev-certs');
  const key = join(dir, 'localhost.key');
  const cert = join(dir, 'localhost.crt');
  if (existsSync(key) && existsSync(cert)) return { key: readFileSync(key), cert: readFileSync(cert) };
  console.warn('\n[MyMath] Chưa có chứng chỉ localhost. Chạy: npm run addin:certs\n');
  return undefined;
}

export default defineConfig(({ mode }) => ({
  base: './',
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  server: { port: 3000, strictPort: true, https: mode === 'addin' ? devHttps() : undefined },
  preview: { port: 3000 },
  build: {
    outDir: 'dist',
    target: 'es2020',
    chunkSizeWarningLimit: 3000,
    rollupOptions: {
      input: {
        main: resolve('index.html'),
        taskpane: resolve('taskpane.html'),
        commands: resolve('commands.html'),
      },
    },
  },
}));
