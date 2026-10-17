import path from 'path';
import { defineConfig, loadEnv, type Plugin, type ViteDevServer } from 'vite';
import react from '@vitejs/plugin-react';

// ローカルの開発サーバーでも /api/gemini（Vercelのサーバー関数）を動かす。
// GeminiのAPIキーはサーバー側だけで使い、ブラウザ用プログラムには入れない。
const devApiGemini = (env: Record<string, string>): Plugin => ({
  name: 'dev-api-gemini',
  configureServer(server: ViteDevServer) {
    if (env.GEMINI_API_KEY) process.env.GEMINI_API_KEY = env.GEMINI_API_KEY;
    if (env.GOOGLE_CLIENT_ID) process.env.GOOGLE_CLIENT_ID = env.GOOGLE_CLIENT_ID;
    server.middlewares.use('/api/gemini', async (req: any, res: any) => {
      const chunks: Buffer[] = [];
      for await (const chunk of req) chunks.push(chunk as Buffer);
      try { req.body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'); } catch { req.body = {}; }
      res.status = (code: number) => { res.statusCode = code; return res; };
      res.json = (data: unknown) => { res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(data)); };
      const mod = await server.ssrLoadModule('/api/gemini.ts');
      await mod.default(req, res);
    });
  },
});

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    return {
      server: {
        port: 3100,
        host: '0.0.0.0',
      },
      plugins: [react(), devApiGemini(env)],
      define: {
        'process.env.GOOGLE_CLIENT_ID': JSON.stringify(env.GOOGLE_CLIENT_ID)
      },
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
        }
      }
    };
});
