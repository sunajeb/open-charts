-- open-charts bridge between PowerPoint's VBA and the engine.
-- PowerPoint runs this with AppleScriptTask("open-charts.scpt", handler, argument), the
-- sandbox-approved way for Office VBA on the Mac to run outside the sandbox. `open-charts addin
-- install` fills in the paths below and compiles this into
-- ~/Library/Application Scripts/com.microsoft.Powerpoint/open-charts.scpt. Installed from the
-- package (open-charts-<version>.pkg), nodePath is empty and cliPath is the engine executable.

property nodePath : "__NODE__"
property cliPath : "__CLI__"
property homeDir : "__HOME__"

-- The add-in server answers only requests carrying the install's token (src/addin/token.ts). The
-- token file holds the header line itself, so curl reads it from the file (-H @file) and the token
-- never appears on a command line.
on tokenHeader()
	return " -H @" & quoted form of (homeDir & "/token")
end tokenHeader

-- One request in, the engine's line-format response out (src/addin/engine.ts, format.ts).
on engine(request)
	set requestText to quoted form of request
	-- 1. The warm engine: the add-in's local server, if it is running (fast path, no node start-up).
	try
		return do shell script "printf '%s' " & requestText & " | /usr/bin/curl -sf --connect-timeout 0.3 --max-time 30 --data-binary @-" & my tokenHeader() & " http://127.0.0.1:4174/api/addin/run" without altering line endings
	end try
	-- 2. Cold: answer with a one-off node process and start the server for the next call.
	try
		set engineCmd to my engineCommand()
		do shell script "nohup " & engineCmd & " addin serve > /dev/null 2>&1 &"
		return do shell script "printf '%s' " & requestText & " | " & engineCmd & " addin run -" without altering line endings
	on error message
		return "OC" & tab & "1" & linefeed & "E" & tab & "open-charts could not start (" & my oneLine(message) & "). Install open-charts again, then restart PowerPoint." & linefeed
	end try
end engine

-- Starts the warm engine in the background unless it is already up. The add-in calls this as it
-- loads, so the first chart inserted gets the fast path above. Returns at once: "OK", plus " verbose"
-- when the user turned verbose logging on (`open-charts addin verbose on`) and " welcome" until the
-- first-run welcome was shown (settings.json "welcomed": true, src/addin/support.ts).
on startServer(ignored)
	try
		set serverCommand to "nohup " & my engineCommand() & " addin serve"
		do shell script "{ /usr/bin/curl -sf --connect-timeout 0.3 --max-time 1 http://127.0.0.1:4174/api/addin/ping || " & serverCommand & "; } > /dev/null 2>&1 &"
		set settingsFile to quoted form of (homeDir & "/settings.json")
		set flags to "OK"
		try
			do shell script "/usr/bin/grep -Eq '\"verbose\": *true' " & settingsFile
			set flags to flags & " verbose"
		end try
		try
			do shell script "/usr/bin/grep -Eq '\"welcomed\": *true' " & settingsFile
		on error
			set flags to flags & " welcome"
		end try
		return flags
	on error message
		return "failed: " & my oneLine(message)
	end try
end startServer

-- The selected elements of an open-charts chart changed (the element mini toolbar): hands the
-- request to the add-in server in the background and returns at once, so PowerPoint never waits.
on elementSelection(request)
	try
		do shell script "printf '%s' " & quoted form of request & " | /usr/bin/curl -sf --connect-timeout 0.3 --max-time 3 --data-binary @-" & my tokenHeader() & " http://127.0.0.1:4174/api/addin/selection > /dev/null 2>&1 &"
	end try
	return ""
end elementSelection

-- Appends one line to addin.log (used by the automation entry points and for diagnostics).
on logLine(lineText)
	try
		do shell script "printf '%s %s\\n' \"$(date '+%Y-%m-%dT%H:%M:%S')\" " & quoted form of my oneLine(lineText) & " >> " & quoted form of (homeDir & "/addin.log")
	end try
	return ""
end logLine

-- "source<tab>destination": copies a file or folder PowerPoint wrote in its sandbox to the
-- destination (used to export a copy of the deck).
on copyOut(arg)
	try
		set AppleScript's text item delimiters to tab
		set {src, dest} to text items of arg
		set AppleScript's text item delimiters to ""
		do shell script "rm -rf " & quoted form of dest & " && mkdir -p \"$(dirname " & quoted form of dest & ")\" && cp -R " & quoted form of src & " " & quoted form of dest & " && rm -rf " & quoted form of src
		return "OK " & dest
	on error message
		set AppleScript's text item delimiters to ""
		return "failed: " & my oneLine(message)
	end try
end copyOut

-- The command that runs the engine: the packaged executable, or node running the CLI source.
on engineCommand()
	if nodePath is "" then return quoted form of cliPath
	return quoted form of my findNode() & " --experimental-strip-types --no-warnings " & quoted form of cliPath
end engineCommand

on findNode()
	try
		do shell script "test -x " & quoted form of nodePath
		return nodePath
	on error
		return do shell script "/bin/zsh -lc 'command -v node'"
	end try
end findNode


on oneLine(txt)
	set AppleScript's text item delimiters to {return, linefeed, tab}
	set parts to text items of (txt as text)
	set AppleScript's text item delimiters to " "
	set joined to parts as text
	set AppleScript's text item delimiters to ""
	return joined
end oneLine
