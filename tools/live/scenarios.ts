// The live scenarios, in the order they run. Each module exports `run(ctx: Ctx)`; it is imported only
// when the scenario runs, so --list and --dry-run load nothing else.
import type { Options, Recorder, ScenarioDef } from './core.ts';
import type { PowerPoint } from './powerpoint.ts';

export const SCENARIOS: ScenarioDef[] = [
  { name: 'install', budgetSec: 180, module: './scenarios/install.ts', explicit: true, description: 'install the add-in (engine, ppam, bridge, pop-up helper), register and load it in PowerPoint' },
  { name: 'insert-all', budgetSec: 120, module: './scenarios/insert-all.ts', description: 'every gallery item on its own slide + content placeholder + "not available"; tags/shapes read back; PNG export' },
  { name: 'edit-apply', budgetSec: 120, module: './scenarios/edit-apply.ts', description: 'Edit data session, Apply, Cancel, ungrouped chart; Apply with PowerPoint closed and after a restart' },
  { name: 'resize', budgetSec: 90, module: './scenarios/resize.ts', description: 'resize re-layout (half, double), moved + Update, too small placeholder and back' },
  { name: 'chart-tools', budgetSec: 120, module: './scenarios/chart-tools.ts', description: 'same scale, totals, CAGR, plain text, reset label, which commands apply, Cmd+Z presses per action' },
  { name: 'directions', budgetSec: 90, module: './scenarios/directions.ts', description: 'direction buttons on selected charts and before an insert, refused direction, ribbon via accessibility' },
  { name: 'agenda', budgetSec: 120, module: './scenarios/agenda.ts', description: 'agenda chapters, nesting, rename, move, split, sorter, hidden slides, table of contents' },
  { name: 'popup', budgetSec: 60, module: './scenarios/popup.ts', description: 'Edit data pop-up: time to show, typed value applied live, Escape puts the chart back' },
  { name: 'handles', budgetSec: 90, module: './scenarios/handles.ts', description: 'drag handles: zoom/window mapping, gap drag, value label pin (moves the mouse)' },
  { name: 'theme', budgetSec: 90, module: './scenarios/theme.ts', description: 'theme colours follow a new colour scheme; same-scale pair re-laid out on Update, Apply and resize' },
  { name: 'fill', budgetSec: 90, module: './scenarios/fill.ts', description: '`open-charts fill` decks open without a repair prompt and export every slide' },
  { name: 'cli-compare', budgetSec: 120, module: './scenarios/cli-compare.ts', description: 'add-in charts vs the CLI writer\'s, rendered by the same PowerPoint, pixel-compared' },
  { name: 'big-deck', budgetSec: 120, module: './scenarios/big-deck.ts', description: 'a 100-slide deck with 50 charts opens, Update chart on the first, middle and last chart, every chart reads back' },
];

/** What a scenario gets: the shared PowerPoint, its output folder, and the recorder. */
export interface Ctx extends Recorder {
  pp: PowerPoint;
  /** out/live/<scenario>, emptied before the scenario runs. */
  out: string;
  /** The repository root. */
  root: string;
  opts: Options;
  /** Results scenarios leave for later ones in the same run (e.g. insert-all's deck for cli-compare). */
  shared: Map<string, unknown>;
}
