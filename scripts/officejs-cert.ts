// Makes the self-signed certificate the Office.js add-in's dev server uses for https://localhost
// (npm run officejs:cert). Office loads add-ins only over trusted https, so the certificate has to
// be trusted once by the machine's owner — this script never touches the keychain; it prints the
// command to run. Kept in ~/Library/Application Support/open-charts/officejs-cert (OPEN_CHARTS_HOME),
// shared by every checkout, so it is trusted once. Re-run with --force to make a new one.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { certDir } from '../src/officejs/server.ts';

const dir = certDir();
const key = join(dir, 'localhost.key'), crt = join(dir, 'localhost.crt'), cnf = join(dir, 'openssl.cnf');

if (!existsSync(crt) || process.argv.includes('--force')) {
  mkdirSync(dir, { recursive: true });
  // A leaf certificate for localhost and 127.0.0.1, valid 825 days (the most macOS accepts for TLS).
  writeFileSync(cnf, `[req]
distinguished_name = dn
x509_extensions = ext
prompt = no
[dn]
CN = open-charts localhost
O = open-charts development
[ext]
basicConstraints = critical, CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost, IP:127.0.0.1
`);
  execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-sha256', '-days', '825', '-keyout', key, '-out', crt, '-config', cnf], { stdio: 'pipe' });
  console.log(`Made ${crt}`);
} else {
  console.log(`Certificate exists: ${crt}`);
}
console.log(`
Trust it once (the add-in will not load in PowerPoint until you do). In Terminal:

  security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "${crt}"

macOS asks for your password. Or double-click the file, and in Keychain Access set
"When using this certificate" to "Always Trust". On Windows: certutil -user -addstore Root localhost.crt
`);
