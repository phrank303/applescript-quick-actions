-- =============================================================
--  Capture One: Export with Collection Hierarchy
--  Version 1.1  (internally known as Attempt #10 ;-)
--
--  Author:       Phrank
--  AI Assistant: Claude (Anthropic) - claude.ai
--  Forum:        Capture One Community
--
--  WHAT IT DOES:
--  Exports the currently selected Project or Group and mirrors
--  the Capture One collection structure as folders on disk.
--  Optionally copies C1 color tags to Finder color labels.
--
--  Example:
--    C1 Collections panel:             Export on disk:
--    Nella & Mad Cold Shooting         Nella & Mad Cold Shooting/
--      Mad Cold                          Mad Cold/
--        Set 01  (11 images)               Set 01/  <- 11 JPEGs
--        Set 07   (7 images)               Set 07/  <-  7 JPEGs
--        Set 09  (16 images)               Set 09/  <- 16 JPEGs
--      Nella                             Nella/
--        Set 08  (54 images)               Set 08/  <- 54 JPEGs
--        Set 11  (66 images)               Set 11/  <- 66 JPEGs
--      Nella & Mad Cold                  Nella & Mad Cold/
--        Set 04  (10 images)               Set 04/  <- 10 JPEGs
--
--  Color label mapping (C1 -> Finder):
--    Red, Orange, Yellow, Green, Blue, Purple, Gray -> matching Finder label
--
--  HOW TO USE:
--  1. Select a Project or Group in the Collections panel.
--  2. Run this script.
--  3. Choose a Process Recipe and a destination folder.
--  4. Wait for Capture One's batch queue to finish.
--  5. Click "Apply Color Labels" to copy C1 tags to Finder labels.
--
--  REQUIREMENTS:
--  - Capture One Pro 23 or later (catalog or session)
--  - At least one Process Recipe must already exist in C1
--  - A Project or Group must be selected in the Collections panel
--
--  INSTALLATION:
--  Place this script in:
--    ~/Library/Scripts/Capture One Scripts/
--  It will then appear under the Scripts menu inside Capture One.
--
--  LICENSE: Free to use and modify. Credit appreciated.
-- =============================================================

use AppleScript version "2.4"
use scripting additions

property gExported : 0
property gErrors : {}
property gRecipeName : ""

-- C1 color tag integer -> Finder label index
-- C1:     0=none 1=red  2=orange 3=yellow 4=green 5=blue 6=purple 7=gray
-- Finder: 0=none 2=red  1=orange 3=yellow 6=green 4=blue 5=purple 7=gray
property gColorMap : {0, 2, 1, 3, 6, 4, 5, 7}

on run
	set gExported to 0
	set gErrors to {}
	set gRecipeName to ""
	
	-- 1. Read the currently selected collection
	set selColl to missing value
	set rootName to ""
	tell application "Capture One"
		try
			set selColl to current collection of current document
			set rootName to name of selColl
		on error errMsg
			tell me to activate
			display dialog "Could not read the current collection." & return & return & errMsg buttons {"OK"} with icon stop
			return
		end try
		set recipeNames to {}
		repeat with r in every recipe of current document
			set end of recipeNames to (name of r) as text
		end repeat
	end tell
	
	if (count of recipeNames) is 0 then
		tell me to activate
		display dialog "No Process Recipes found." & return & return & "Please create at least one recipe in the Output panel first." buttons {"OK"} with icon stop
		return
	end if
	
	-- 2. Choose recipe
	tell me to activate
	set recipeChoice to choose from list recipeNames Â
		with prompt "Selected collection:  " & rootName & return & return & "Which Process Recipe should be used for the export?" Â
		with title "CO Export with Collection Hierarchy"
	if recipeChoice is false then return
	set gRecipeName to item 1 of recipeChoice
	
	-- 3. Choose destination folder
	tell me to activate
	set destHFS to choose folder with prompt "Choose the export destination folder for:  " & rootName
	set destBase to POSIX path of destHFS
	if destBase does not end with "/" then set destBase to destBase & "/"
	
	-- 4. Save original document output + recipe settings
	set origDocOutput to missing value
	set origRecipeRootType to missing value
	set origRecipeSubFolder to ""
	tell application "Capture One"
		tell current document
			try
				set origDocOutput to output
			end try
			set r to recipe gRecipeName
			try
				set origRecipeRootType to root folder type of r
			end try
			try
				set origRecipeSubFolder to output sub folder of r
			end try
			set root folder type of r to output location
			set output sub folder of r to ""
		end tell
	end tell
	
	-- 5. Record start time (used later to find freshly exported files)
	set exportStartTime to current date
	
	-- 6. Create root folder and walk the collection tree
	set rootPath to destBase & (my safeName(rootName)) & "/"
	do shell script "mkdir -p " & quoted form of rootPath
	my exportColl(selColl, rootPath)
	
	-- 7. Restore original document output + recipe settings
	tell application "Capture One"
		tell current document
			try
				if origDocOutput is not missing value then
					set output to origDocOutput
				end if
			end try
			set r to recipe gRecipeName
			try
				set root folder type of r to origRecipeRootType
			end try
			try
				set output sub folder of r to origRecipeSubFolder
			end try
		end tell
	end tell
	
	-- 8. Summary dialog
	tell me to activate
	set msg to "Export queued successfully!" & return & return
	set msg to msg & "Collection:   " & rootName & return
	set msg to msg & "Recipe:       " & gRecipeName & return
	set msg to msg & "Queued:       " & gExported & " images" & return
	set msg to msg & "Destination:  " & rootPath & return & return
	set msg to msg & "Wait for Capture One's batch queue to finish," & return
	set msg to msg & "then click 'Apply Color Labels'."
	
	if (count of gErrors) > 0 then
		set msg to msg & return & return & (count of gErrors) & " album(s) had errors:"
		repeat with e in gErrors
			set msg to msg & return & "  - " & e
		end repeat
	end if
	
	tell me to activate
	set btn to button returned of (display dialog msg Â
		buttons {"Open Destination", "Apply Color Labels", "Done"} default button "Apply Color Labels" Â
		with title "CO Export with Collection Hierarchy")
	
	if btn is "Open Destination" then
		tell application "Finder"
			open (POSIX file rootPath as alias)
			activate
		end tell
		
	else if btn is "Apply Color Labels" then
		-- 9. Confirm batch queue is done before applying labels
		tell me to activate
		set confirmed to button returned of (display dialog Â
			"Has Capture One finished processing the batch queue?" & return & return & Â
			"(Check the queue indicator in the top bar is empty.)" Â
			buttons {"Cancel", "Yes, apply labels"} default button "Yes, apply labels" Â
			with title "CO Export with Collection Hierarchy")
		if confirmed is "Yes, apply labels" then
			my applyColorLabels(selColl, rootPath, exportStartTime)
		end if
	end if
end run


-- =============================================================
--  exportColl  -  walks the C1 collection tree recursively
-- =============================================================
on exportColl(coll, basePath)
	set kids to {}
	tell application "Capture One"
		try
			set kids to collections of coll
		end try
	end tell
	
	if (count of kids) > 0 then
		repeat with kid in kids
			set kidName to ""
			tell application "Capture One"
				set kidName to name of kid
			end tell
			set kidPath to basePath & (my safeName(kidName)) & "/"
			do shell script "mkdir -p " & quoted form of kidPath
			my exportColl(kid, kidPath)
		end repeat
		
	else
		set theVariants to {}
		set collName to ""
		tell application "Capture One"
			try
				set theVariants to every variant of coll
			end try
			set collName to name of coll
		end tell
		if (count of theVariants) is 0 then return
		
		try
			set destFileSpec to POSIX file basePath
			tell application "Capture One"
				tell current document
					set output to destFileSpec
				end tell
				set jobResult to process theVariants recipe gRecipeName
			end tell
			if (jobResult as text) starts with "ERROR" then
				set gErrors to gErrors & {collName & ": " & jobResult}
			else
				set gExported to gExported + (count of theVariants)
			end if
		on error errExp
			set gErrors to gErrors & {collName & ": " & errExp}
		end try
	end if
end exportColl


-- =============================================================
--  applyColorLabels  -  recursively walks collection tree,
--  reads each variant's color tag and output events,
--  then sets the matching Finder label on the exported file
-- =============================================================
on applyColorLabels(coll, basePath, startTime)
	set kids to {}
	tell application "Capture One"
		try
			set kids to collections of coll
		end try
	end tell
	
	if (count of kids) > 0 then
		repeat with kid in kids
			set kidName to ""
			tell application "Capture One"
				set kidName to name of kid
			end tell
			set kidPath to basePath & (my safeName(kidName)) & "/"
			my applyColorLabels(kid, kidPath, startTime)
		end repeat
		
	else
		set theVariants to {}
		tell application "Capture One"
			try
				set theVariants to every variant of coll
			end try
		end tell
		if (count of theVariants) is 0 then return
		
		repeat with v in theVariants
			set ctag to 0
			set outputPath to ""
			
			tell application "Capture One"
				try
					set ctag to color tag of v
				end try
				-- Find the most recent output event for this variant
				try
					set allEvents to every output event of v
					repeat with ev in allEvents
						try
							if (date of ev) >= startTime and (exists of ev) then
								set outputPath to path of ev
							end if
						end try
					end repeat
				end try
			end tell
			
			-- Apply Finder label if the file exists and has a color tag
			if outputPath is not "" and ctag > 0 then
				try
					-- Map C1 color tag to Finder label index
					set finderLabel to item (ctag + 1) of gColorMap
					set outputAlias to (POSIX file outputPath) as alias
					tell application "Finder"
						set label index of outputAlias to finderLabel
					end tell
				end try
			end if
		end repeat
	end if
end applyColorLabels


-- =============================================================
--  safeName  -  replaces "/" in collection names with "-"
-- =============================================================
on safeName(n)
	set cleaned to ""
	repeat with c in (characters of n)
		if (c as text) is "/" then
			set cleaned to cleaned & "-"
		else
			set cleaned to cleaned & (c as text)
		end if
	end repeat
	return cleaned
end safeName
