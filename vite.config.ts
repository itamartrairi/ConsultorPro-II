import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig, loadEnv} from 'vite';

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  
  // Find key in env or process.env with potential variations and spelling (e.g. Gemini_api_kei)
  const lookupKey = () => {
    const keysToCheck = [
      'GEMINI_API_KEY',
      'Gemini_api_kei',
      'GEMINI_API_KEI',
      'gemini_api_kei',
      'gemini_api_key',
      'VITE_GEMINI_API_KEY',
      'VITE_GEMINI_API_KEI'
    ];
    // Check direct match
    for (const k of keysToCheck) {
      if (env[k]) return env[k];
      if (process.env[k]) return process.env[k];
    }
    // Check case-insensitive match
    const pooled = { ...process.env, ...env };
    for (const key of Object.keys(pooled)) {
      const l = key.toLowerCase();
      if (
        l === 'gemini_api_kei' || 
        l === 'gemini_api_key' || 
        l === 'gemini_key' || 
        l === 'gemini_kei' || 
        l.includes('gemini_api') || 
        l.includes('gemini_key') || 
        l.includes('gemini_kei')
      ) {
        const val = pooled[key];
        if (val && val !== 'undefined' && val !== 'null') return val;
      }
    }
    return "";
  };

  const resolvedApiKey = lookupKey().trim();

  return {
    base: './',
    plugins: [react(), tailwindcss()],
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(resolvedApiKey),
      'import.meta.env.VITE_GEMINI_API_KEY': JSON.stringify(resolvedApiKey),
      'process.env.VITE_GEMINI_API_KEY': JSON.stringify(resolvedApiKey),
    },
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify - file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
    },
  };
});
