-- Drag-and-drop droplet for capture.py. Drop one or more .pptx files onto
-- the built app; each gets captured into a "<name>_frames" folder next to
-- the source file. Build with: osacompile -o DropToCapture.app DropToCapture.applescript
--
-- Hands the actual work off to Terminal rather than running screencapture
-- itself: an ad-hoc-signed osacompile app has no stable code-signing
-- identity, so macOS can re-prompt for Screen Recording on this app even
-- after it's been granted. Terminal already holds that permission, so this
-- sidesteps the problem entirely instead of fighting TCC.

property repoDir : (POSIX path of (path to home folder)) & "ppt-de-animator"

on run
	display dialog "Drop a .pptx file onto this app to capture it." buttons {"OK"} default button 1 with icon note
end run

on open theFiles
	repeat with f in theFiles
		set deckPath to POSIX path of f
		if deckPath ends with ".pptx" then
			processDeck(deckPath)
		else
			display notification "Skipped (not a .pptx): " & deckPath with title "PPT De-Animator"
		end if
	end repeat
end open

on processDeck(deckPath)
	set shellCmd to "" & ¬
		"deck=" & quoted form of deckPath & "; " & ¬
		"dir=$(dirname \"$deck\"); " & ¬
		"base=$(basename \"$deck\" .pptx); " & ¬
		"out=\"$dir/${base}_frames\"; " & ¬
		"cd " & quoted form of repoDir & " && python3 capture.py \"$deck\" \"$out\"; " & ¬
		"osascript -e 'display notification \"" & ¬
		"Finished, check the _frames folder next to the source file.\" with title \"PPT De-Animator\"'"

	tell application "Terminal"
		activate
		do script shellCmd
	end tell
end processDeck
