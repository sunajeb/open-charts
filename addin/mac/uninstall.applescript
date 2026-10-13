-- Uninstall open-charts.app (the open-charts package puts it in ~/Applications): removes the add-in,
-- its engine, the pop-up helper, the bridge script and the open-charts command, with
-- `open-charts addin uninstall` run by the installed engine.
on run
	set engine to (POSIX path of (path to home folder)) & "Library/Application Support/open-charts/engine/open-charts"
	try
		do shell script "test -x " & quoted form of engine
	on error
		display dialog "open-charts is not installed (or was already removed)." buttons {"OK"} default button "OK" with title "open-charts"
		return
	end try
	display dialog "Remove the open-charts add-in from PowerPoint?" & return & return & "Charts already in your presentations stay as they are." buttons {"Cancel", "Remove"} default button "Remove" cancel button "Cancel" with title "open-charts" with icon caution
	try
		do shell script quoted form of engine & " addin uninstall"
		display dialog "open-charts was removed." buttons {"OK"} default button "OK" with title "open-charts"
	on error message
		display dialog "open-charts could not be removed: " & message buttons {"OK"} default button "OK" with title "open-charts" with icon stop
	end try
end run
