-- open-charts bridge between PowerPoint's VBA and the engine.
-- PowerPoint runs this with AppleScriptTask("open-charts.scpt", handler, argument), the
-- sandbox-approved way for Office VBA on the Mac to run outside the sandbox. `open-charts addin
-- install` fills in the paths below and compiles this into
-- ~/Library/Application Scripts/com.microsoft.Powerpoint/open-charts.scpt.

property nodePath : "__NODE__"
property cliPath : "__CLI__"
property homeDir : "__HOME__"

-- One request in, the engine's line-format response out (src/addin/engine.ts, format.ts).
on engine(request)
	set requestText to quoted form of request
	-- 1. The warm engine: the add-in's local server, if it is running (fast path, no node start-up).
	try
		return do shell script "printf '%s' " & requestText & " | /usr/bin/curl -sf --connect-timeout 0.3 --max-time 30 --data-binary @- http://127.0.0.1:4174/api/addin/run" without altering line endings
	end try
	-- 2. Cold: answer with a one-off node process and start the server for the next call.
	try
		set nodeBin to quoted form of my findNode()
		set engineArgs to " --experimental-strip-types --no-warnings " & quoted form of cliPath & " addin "
		do shell script "nohup " & nodeBin & engineArgs & "serve > /dev/null 2>&1 &"
		return do shell script "printf '%s' " & requestText & " | " & nodeBin & engineArgs & "run -" without altering line endings
	on error message
		return "OC" & tab & "1" & linefeed & "E" & tab & "The open-charts engine failed: " & my oneLine(message) & linefeed
	end try
end engine

-- Starts the warm engine in the background unless it is already up. The add-in calls this as it
-- loads, so the first chart inserted gets the fast path above. Returns at once.
on startServer(ignored)
	try
		set serverCommand to "nohup " & quoted form of my findNode() & " --experimental-strip-types --no-warnings " & quoted form of cliPath & " addin serve"
		do shell script "{ /usr/bin/curl -sf --connect-timeout 0.3 --max-time 1 http://127.0.0.1:4174/api/addin/ping || " & serverCommand & "; } > /dev/null 2>&1 &"
		return "OK"
	on error message
		return "failed: " & my oneLine(message)
	end try
end startServer

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
