// Who may talk to open-charts' local servers (the editor, the add-in server, the Office.js dev host).
// They listen on 127.0.0.1 only, but any web page open in the user's browser can still send requests
// there, so each request is checked (docs/ux-editor.md "Security"):
//   - Host must be this server (127.0.0.1 or localhost, its port): a page that re-points its own
//     domain at 127.0.0.1 (DNS rebinding) sends its domain as Host and is refused.
//   - Origin, when the browser sends one, must be this server: another site's fetch or form post is
//     refused. No CORS headers are sent, so other sites cannot read answers either.
//   - With a token (the add-in server), every route but the public ones needs the X-OC-Token header.
import type { IncomingMessage } from 'node:http';
import { sameToken, TOKEN_HEADER } from './addin/token.ts';

export interface Guard {
  /** The port the server listens on (as it is known once listening). */
  port: () => number;
  scheme?: 'http' | 'https';
  /** Required on every non-public request when set. */
  token?: string;
  /** Requests that need no token (static pages, liveness). */
  isPublic?: (req: IncomingMessage) => boolean;
  /** More names this server may be reached by (Office on the web through a tunnel). */
  hosts?: string[];
}

/** Why `req` is refused, or undefined when it may proceed. */
export function refusal(req: IncomingMessage, g: Guard): string | undefined {
  const port = g.port();
  const names = ['127.0.0.1', 'localhost', ...(g.hosts ?? [])].map((h) => h.toLowerCase());
  const host = (req.headers.host ?? '').toLowerCase();
  const own = names.flatMap((n) => [`${n}:${port}`, ...(g.hosts?.includes(n) ? [n] : [])]);
  if (!own.includes(host)) return 'wrong Host';
  const origin = req.headers.origin;
  if (origin !== undefined) {
    const scheme = g.scheme ?? 'http';
    if (!own.some((h) => origin.toLowerCase() === `${scheme}://${h}`)) return 'foreign Origin';
  }
  if (g.token && !(g.isPublic?.(req) ?? false)) {
    const presented = req.headers[TOKEN_HEADER.toLowerCase()];
    if (!sameToken(Array.isArray(presented) ? presented[0] : presented, g.token)) return 'missing or wrong token';
  }
  return undefined;
}
