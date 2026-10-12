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
	try
		-- The request goes in on stdin (no temp file: file I/O from this context fails with -50).
		set command to "printf '%s' " & quoted form of request & " | " & quoted form of my findNode() & " --experimental-strip-types --no-warnings " & quoted form of cliPath & " addin run -"
		return do shell script command without altering line endings
	on error message
		return "OC" & tab & "1" & linefeed & "E" & tab & "The open-charts engine failed: " & my oneLine(message) & linefeed
	end try
end engine

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
