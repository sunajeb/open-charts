// Stand-in for the PowerPoint add-in's server in the editor E2E (`node host.ts <port>`): it confirms
// every Apply as the add-in does (src/addin/server.ts), through the same onSession hook, with the
// outcome picked by the session id's prefix:
//   perm…  PowerPoint does not answer (macOS permission prompt), error after 3 s
//   gone…  PowerPoint is not running, error at once
//   other  the chart is drawn after 300 ms
import { serve, setSession } from '../../src/server.ts';
import { applyErrorMessage } from '../../src/addin/server.ts';

const port = Number(process.argv[2]);
serve(port, {}, {
  confirmsApply: true,
  onSession: (s) => {
    if (s.state !== 'applied') return;
    const later = (ms: number, body: { state: string; message?: string }) => setTimeout(() => setSession(s.id, body), ms);
    if (s.id.startsWith('perm')) later(3000, { state: 'error', message: applyErrorMessage({ message: 'timed out', killed: true }) });
    else if (s.id.startsWith('gone')) later(0, { state: 'error', message: applyErrorMessage({ message: 'execution error: PowerPoint is not running. (-600)' }) });
    else later(300, { state: 'drawn' });
  },
});
