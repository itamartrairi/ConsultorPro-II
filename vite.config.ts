import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig, loadEnv} from 'vite';

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  
  // A chave do Gemini NUNCA é embutida no JavaScript do navegador (qualquer visitante
  // conseguiria lê-la). Ela fica só no servidor: variável GEMINI_API_KEY da Netlify
  // Function /api/gemini/* (ou do server.ts no Docker/dev).
  void env;

  return {
    base: './',
    plugins: [react(), tailwindcss()],
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(''),
      'import.meta.env.VITE_GEMINI_API_KEY': JSON.stringify(''),
      'process.env.VITE_GEMINI_API_KEY': JSON.stringify(''),
    },
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    build: {
      // Separa as bibliotecas grandes em arquivos próprios: carregam em paralelo e
      // ficam em cache do navegador entre as atualizações do app.
      chunkSizeWarningLimit: 1500,
      rollupOptions: {
        output: {
          manualChunks(id: string) {
            if (!id.includes('node_modules')) return;
            if (/[\\/]node_modules[\\/](react|react-dom|scheduler|react-router|react-router-dom)[\\/]/.test(id)) return 'vendor-react';
            if (id.includes('/@firebase/') || id.includes('/firebase/')) return 'vendor-firebase';
            if (id.includes('/recharts/') || id.includes('/d3-') || id.includes('/victory-vendor/')) return 'vendor-charts';
            if (id.includes('/xlsx/')) return 'vendor-xlsx';
            if (id.includes('/jspdf') || id.includes('/html2canvas') || id.includes('/canvg/') || id.includes('/dompurify/')) return 'vendor-pdf';
            if (id.includes('/@google/genai/')) return 'vendor-genai';
            if (id.includes('/motion/') || id.includes('/framer-motion/') || id.includes('/motion-dom/')) return 'vendor-motion';
            if (id.includes('/lucide-react/')) return 'vendor-icons';
          },
        },
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
    },
  };
});
