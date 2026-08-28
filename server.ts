import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
import { dbRouter } from "./src/db/routes";

dotenv.config();

function getSystemGeminiApiKey(): string | undefined {
  if (process.env.GEMINI_API_KEY && process.env.GEMINI_API_KEY !== 'undefined' && process.env.GEMINI_API_KEY !== 'null' && process.env.GEMINI_API_KEY.trim() !== '') {
    return process.env.GEMINI_API_KEY.trim();
  }
  if (process.env.VITE_GEMINI_API_KEY && process.env.VITE_GEMINI_API_KEY !== 'undefined' && process.env.VITE_GEMINI_API_KEY !== 'null' && process.env.VITE_GEMINI_API_KEY.trim() !== '') {
    return process.env.VITE_GEMINI_API_KEY.trim();
  }

  // Look through all environment keys
  for (const envKey of Object.keys(process.env)) {
    const lKey = envKey.toLowerCase();
    // Match any version of gemini key spelling, such as Gemini_api_key, etc.
    if (
      lKey === 'gemini_api_key' ||
      lKey === 'vite_gemini_api_key' ||
      lKey === 'gemini_key' ||
      lKey === 'google_api_key' ||
      lKey.includes('gemini') ||
      lKey.includes('genai')
    ) {
      const val = process.env[envKey];
      if (typeof val === 'string' && val.trim().length > 10 && val !== 'undefined' && val !== 'null') {
        console.log(`[Gemini Config] Detected API key under environment variable: ${envKey}`);
        return val.trim();
      }
    }
  }
  return undefined;
}

async function startServer() {
  console.log("Environment Keys on Server Startup:", Object.keys(process.env).filter(k => k.toLowerCase().includes("gemini") || k.toLowerCase().includes("key") || k.toLowerCase().includes("kei")));
  const app = reportExpressErrorsAndStart();
  const PORT = 3000;

  function reportExpressErrorsAndStart() {
    return express();
  }

  app.use(express.json());

  // Netlify Database / PostgreSQL API
  app.use("/api/db", dbRouter);

  // Check system API Key status
  app.get("/api/gemini/config", (req, res) => {
    const sysKey = getSystemGeminiApiKey();
    const hasKey = typeof sysKey === 'string' && 
                   sysKey.trim() !== '' && 
                   sysKey !== 'undefined' &&
                   sysKey !== 'null';
    
    // Find matching keys in process.env for debugging info
    const matchedKeys = Object.keys(process.env).filter(k => {
      const l = k.toLowerCase();
      return l.includes('gemini') || l.includes('key') || l.includes('kei');
    });

    res.json({ 
      hasSystemKey: hasKey, 
      detectedKeys: matchedKeys,
      snippet: (hasKey && sysKey) ? (sysKey.substring(0, 4) + "..." + sysKey.substring(sysKey.length - 4)) : null,
      length: (hasKey && sysKey) ? sysKey.length : 0
    });
  });

  // Test Gemini connection endpoint
  app.post("/api/gemini/test", async (req, res) => {
    try {
      const rawCustomKey = req.headers['x-custom-api-key'] || req.headers['X-Custom-Api-Key'];
      const systemKey = getSystemGeminiApiKey();
      let apiKey = systemKey;
      if (typeof rawCustomKey === 'string' && rawCustomKey.trim().length > 10) {
        apiKey = rawCustomKey.trim();
      }

      if (!apiKey || apiKey === "undefined" || apiKey.trim() === "") {
        return res.status(400).json({ 
          ok: false, 
          error: "Chave API do Gemini não configurada. Por favor, insira sua GEMINI_API_KEY no menu Configurações." 
        });
      }

      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });

      const testModels = ["gemini-3.1-flash-lite", "gemini-3.7-flash", "gemini-flash-latest"];
      let workingModel: string | null = null;
      let lastErr: any = null;

      for (const m of testModels) {
        try {
          const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error("Timeout de teste da API")), 10000)
          );
          const r: any = await Promise.race([
            ai.models.generateContent({
              model: m,
              contents: "Responda apenas 'OK'."
            }),
            timeoutPromise
          ]);
          if (r && (r.text || r.candidates?.[0]?.content?.parts?.[0]?.text)) {
            workingModel = m;
            break;
          }
        } catch (e: any) {
          console.warn(`[Gemini Test] Model ${m} attempt failed:`, e?.message || e);
          lastErr = e;
        }
      }

      if (workingModel) {
        return res.json({ 
          ok: true, 
          model: workingModel, 
          message: `Conexão bem-sucedida com a IA do Google Gemini (modelo: ${workingModel})!` 
        });
      } else {
        const rawErrMsg = lastErr?.message || "";
        let friendlyErr = rawErrMsg;
        if (rawErrMsg.includes("API key not valid") || rawErrMsg.includes("API_KEY_INVALID")) {
          friendlyErr = "Chave de API inválida. Verifique se copiou a chave completa gerada no Google AI Studio.";
        }
        return res.status(500).json({ 
          ok: false, 
          error: friendlyErr || "Falha ao conectar aos modelos do Gemini. Verifique sua chave de API." 
        });
      }
    } catch (e: any) {
      return res.status(500).json({ ok: false, error: e?.message || "Erro inesperado ao testar Gemini." });
    }
  });

  // API Route for Gemini
  app.post("/api/gemini/generate", async (req, res) => {
    try {
      const { model, contents, config } = req.body;
      
      const rawCustomKey = req.headers['x-custom-api-key'] || req.headers['X-Custom-Api-Key'];
      const systemKey = getSystemGeminiApiKey();
      let apiKey = systemKey;
      if (typeof rawCustomKey === 'string' && rawCustomKey.trim().length > 10) {
        apiKey = rawCustomKey.trim();
      }

      if (!apiKey || apiKey === "undefined" || apiKey.trim() === "") {
        console.warn("[Gemini Server] Warning: No API Key configured!");
        return res.status(400).json({ error: "Chave API do Gemini não configurada. Por favor, adicione sua chave no menu Configurações." });
      }

      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });

      const GEMINI_PRIMARY = "gemini-3.1-flash-lite";
      const GEMINI_CHAIN = [
        "gemini-3.1-flash-lite",
        "gemini-3.7-flash",
        "gemini-flash-latest"
      ];

      function normalizeServerModel(m?: string): string {
        if (!m) return GEMINI_PRIMARY;
        if (m.includes("2.5") || m.includes("2.0") || m.includes("1.5") || m.includes("1.0") || m.includes("3.6")) {
          return "gemini-3.1-flash-lite";
        }
        return m;
      }

      const initialModel = normalizeServerModel(model);
      const candidateModels = Array.from(new Set([initialModel, ...GEMINI_CHAIN]));

      let response: any = null;
      let lastErr: any = null;

      for (const modelToTry of candidateModels) {
        try {
          const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error("Tempo limite de resposta do Gemini excedido (15s)")), 15000)
          );
          response = await Promise.race([
            ai.models.generateContent({
              model: modelToTry,
              contents: contents,
              config: config
            }),
            timeoutPromise
          ]);
          if (response && response.text !== undefined) {
            break;
          }
        } catch (err: any) {
          console.warn(`[Gemini Proxy] Model ${modelToTry} failed:`, err?.message || err);
          lastErr = err;
        }
      }

      if (!response || response.text === undefined) {
        const errMsg = lastErr?.message || "Não foi possível conectar aos modelos do Gemini no momento. Tente novamente mais tarde.";
        return res.status(500).json({ error: errMsg });
      }

      res.json({ text: response.text });
    } catch (error: any) {
      console.error("Erro no proxy do Gemini:", error);
      res.status(500).json({ error: error.message || "Erro na geração de conteúdo da IA." });
    }
  });

  // Serve static/vite assets
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
