// Usage text only, so `open-charts --help` doesn't load the add-in modules.
export const ADDIN_USAGE = `Usage: open-charts addin build [-o dist/open-charts.ppam]
       open-charts addin install [--no-register]
       open-charts addin uninstall
       open-charts addin types
       open-charts addin verbose on|off        (detailed timings in addin.log; off: results and errors)
       open-charts addin update [on|off]       (check for a new release now; turn the daily check on or off)
       open-charts addin run <request.txt|->   (called by the add-in)
       open-charts addin serve                 (edit-data server, started by the add-in)
       open-charts addin register              (Mac package install: run when PowerPoint starts)`;
