// The add-in server's per-install secret (docs/ux-editor.md "Security"). The server on 127.0.0.1:4174
// drives PowerPoint (Apply runs a macro) and holds the charts being edited, so it answers only callers
// that can read this file: 32 random bytes (hex) in the add-in's home, readable by the user alone.
//   - Node (the server, `addin run`, the live test tools): the X-OC-Token header (tokenHeaders).
//   - The AppleScript bridge (Mac VBA) and the Windows VBA read the file and send the same header.
//   - Pages the server opens (editor pop-up, toolbar, browser fallback): the URL fragment
//     #oc-token=… (never the query string: a fragment is not sent to the server, logged or put in a
//     Referer); the page moves it into the header of its own requests and out of the address bar.
import { randomBytes, timingSafeEqual } from 'node:crypto';
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { addinHome } from './launcher.ts';

export const TOKEN_HEADER = 'X-OC-Token';
export const TOKEN_FRAGMENT = 'oc-token';

/**
 * <home>/token holds the header line itself ("X-OC-Token: <64 hex>"), so the Mac bridge hands the
 * file to curl (`-H @file`) and the token never appears on a command line (`ps` shows those).
 */
export const tokenPath = (home = addinHome()) => join(home, 'token');
const TOKEN_LINE = new RegExp(`^\\uFEFF?${TOKEN_HEADER}: ([0-9a-f]{64})\\s*$`);

/** The install's token, or undefined when there is none (not installed) or the file is not one. */
export function readToken(home = addinHome()): string | undefined {
  try {
    return TOKEN_LINE.exec(readFileSync(tokenPath(home), 'utf8'))?.[1];
  } catch { return undefined; }
}

/** The install's token, created (user-only permissions) when missing or damaged. */
export function ensureToken(home = addinHome()): string {
  const file = tokenPath(home);
  const have = readToken(home);
  if (have) {
    try { chmodSync(file, 0o600); } catch { /* Windows: %LOCALAPPDATA% is per-user already */ }
    return have;
  }
  mkdirSync(home, { recursive: true });
  const t = randomBytes(32).toString('hex');
  writeFileSync(file, `${TOKEN_HEADER}: ${t}\n`, { mode: 0o600 });
  try { chmodSync(file, 0o600); } catch { /* as above */ }
  return t;
}

/** Headers for a call to the add-in server from Node. */
export const tokenHeaders = (home = addinHome()): Record<string, string> => {
  const t = readToken(home);
  return t ? { [TOKEN_HEADER]: t } : {};
};

/** A page URL carrying the token in its fragment (unchanged without a token). */
export const withToken = (url: string, token: string | undefined) => (token ? `${url}#${TOKEN_FRAGMENT}=${token}` : url);

/** Constant-time comparison of a presented token with the expected one. */
export function sameToken(presented: string | undefined, expected: string): boolean {
  if (typeof presented !== 'string') return false;
  const a = Buffer.from(presented), b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

/**
 * The script a page served by the add-in server runs first (editor, toolbar): takes the token out of
 * the URL fragment (kept for this tab in sessionStorage, so a reload works) and out of the address
 * bar and history, and adds it as the header of the page's own same-origin fetches.
 */
export const TOKEN_BOOT = `<script>(function () {
  var m = /(?:^#|&)${TOKEN_FRAGMENT}=([0-9a-f]{64})(?:&|$)/.exec(location.hash), t = m ? m[1] : null;
  try { if (t) sessionStorage.setItem('${TOKEN_FRAGMENT}', t); else t = sessionStorage.getItem('${TOKEN_FRAGMENT}'); } catch (e) {}
  if (m) history.replaceState(history.state, '', location.pathname + location.search);
  if (!t) return;
  var f = window.fetch;
  window.fetch = function (input, init) {
    var url = new URL(typeof input === 'string' || input instanceof URL ? String(input) : input.url, location.href);
    if (url.origin !== location.origin) return f.call(this, input, init);
    init = Object.assign({}, init);
    var h = new Headers(init.headers || (input instanceof Request ? input.headers : undefined));
    h.set('${TOKEN_HEADER}', t);
    init.headers = h;
    return f.call(this, input, init);
  };
})();</script>
`;

/** `html` with TOKEN_BOOT before its first script. */
export function withTokenBoot(html: string): string {
  const at = html.indexOf('<script');
  return at < 0 ? html + TOKEN_BOOT : html.slice(0, at) + TOKEN_BOOT + html.slice(at);
}
