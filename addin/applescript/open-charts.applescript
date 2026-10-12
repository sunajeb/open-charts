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
		set requestFile to homeDir & "/request.txt"
		my writeUtf8(requestFile, request)
		set command to quoted form of my findNode() & " --experimental-strip-types --no-warnings " & quoted form of cliPath & " addin run " & quoted form of requestFile
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

on findNode()
	try
		do shell script "test -x " & quoted form of nodePath
		return nodePath
	on error
		return do shell script "/bin/zsh -lc 'command -v node'"
	end try
end findNode

on writeUtf8(filePath, txt)
	set fh to open for access (POSIX file filePath) with write permission
	try
		set eof fh to 0
		write txt to fh as «class utf8»
		close access fh
	on error message
		close access fh
		error message
	end try
end writeUtf8

on oneLine(txt)
	set AppleScript's text item delimiters to {return, linefeed, tab}
	set parts to text items of (txt as text)
	set AppleScript's text item delimiters to " "
	set joined to parts as text
	set AppleScript's text item delimiters to ""
	return joined
end oneLine
