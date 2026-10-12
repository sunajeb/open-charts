import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { readFileSync } from 'node:fs';
import { createMeasurer } from './fonts.ts';
import { buildDeck, type DeckSpec } from './deck.ts';

const EDITOR = new URL('./web/editor.html', import.meta.url);

async function body(req: IncomingMessage): Promise<DeckSpec> {
  const chunks: Buffer[] = [];
  for await (const c of req) chunks.push(c as Buffer);
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

/** Extra routes (the PowerPoint add-in's edit session); returns true when it handled the request. */
export type Route = (req: IncomingMessage, res: ServerResponse) => boolean | Promise<boolean>;

/** Local editor: datasheet + live preview, renders with the same engine as the CLI. */
export function serve(port: number, fonts: { font?: string; bold?: string } = {}, extra?: Route) {
  const m = createMeasurer(fonts.font, fonts.bold);
  const server = createServer(async (req, res) => {
    try {
      if (extra && await extra(req, res)) return;
      if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
        res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
        res.end(readFileSync(EDITOR));
        return;
      }
      if (req.method === 'POST' && req.url === '/api/render') {
        const { slides } = buildDeck(await body(req), m);
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ slides: slides.map((s) => ({ svg: s.svg, issues: s.issues })) }));
        return;
      }
      if (req.method === 'POST' && req.url === '/api/pptx') {
        const { pptx } = buildDeck(await body(req), m);
        res.writeHead(200, {
          'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'content-disposition': 'attachment; filename="charts.pptx"',
        });
        res.end(Buffer.from(pptx));
        return;
      }
      res.writeHead(404).end('Not found');
    } catch (err) {
      res.writeHead(400, { 'content-type': 'application/json' });
      res.end(JSON.stringify({ error: (err as Error).message }));
    }
  });
  server.listen(port, '127.0.0.1', () => console.log(`open-charts editor: http://localhost:${port}`));
  return server;
}
