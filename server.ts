import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import { dbRouter } from "./src/db/routes";
import { handleKiwifyWebhook } from "./src/server/kiwifyWebhook";
import { geminiConfig, geminiGenerate, geminiTest } from "./src/server/geminiCore";
import { verifyFirebaseUser } from "./src/server/firebaseAdmin";

dotenv.config();

async function startServer() {
  const app = reportExpressErrorsAndStart();
  const PORT = Number(process.env.PORT) || 3000;

  function reportExpressErrorsAndStart() {
    return express();
  }

  app.use(express.json({ limit: '10mb' }));

  // Netlify Database / PostgreSQL API
  app.use("/api/db", dbRouter);

  // Kiwify — libera automaticamente o plano comprado (Mensal/Anual) quando o
  // pagamento é aprovado. Veja instruções de configuração em src/server/kiwifyWebhook.ts
  app.post("/api/webhooks/kiwify", handleKiwifyWebhook);

  // --- Proxy do Gemini (mesma lógica da Netlify Function netlify/functions/gemini.mts) ---
  app.get("/api/gemini/config", (_req, res) => {
    const r = geminiConfig();
    res.status(r.status).json(r.body);
  });

  // Exige login: sem isso, qualquer pessoa usaria a chave do sistema pelo seu servidor.
  const requireLogin: express.RequestHandler = async (req, res, next) => {
    const user = await verifyFirebaseUser(req.headers.authorization);
    if (!user) return res.status(401).json({ ok: false, error: "Faça login novamente para usar a IA." });
    next();
  };

  app.post("/api/gemini/test", requireLogin, async (req, res) => {
    const r = await geminiTest(req.headers["x-custom-api-key"]);
    res.status(r.status).json(r.body);
  });

  app.post("/api/gemini/generate", requireLogin, async (req, res) => {
    const r = await geminiGenerate(req.body, req.headers["x-custom-api-key"]);
    res.status(r.status).json(r.body);
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
