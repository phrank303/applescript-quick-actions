-- =============================================================
--  Capture One: Create Album from Text File
--  Version 1.4
--
--  Author:       Phrank
--  AI Assistant: Claude (Anthropic) - claude.ai
--  GitHub:       https://github.com/phrank303/applescript-quick-actions
--
--  HOW TO USE:
--  1. In Capture One, click the folder or album you want to search.
--  2. Run this script.
--  3. Select the text file with filenames (one per line).
--  4. Enter album name, optionally set rating and color tag.
--
--  TIPS:
--  - Filenames are matched case-sensitively.
--  - Extensions in the text file are optional.
--  - Select a specific collection, not "All Images", for speed.
--
--  REQUIREMENTS:
--  - Capture One Pro 23 or later, catalog or session open,
--    a collection selected before running.
--
--  INSTALLATION:
--  ~/Library/Scripts/Capture One Scripts/
--
--  LICENSE: MIT - For personal use only. Credit appreciated.
-- =============================================================

tell application "/Applications/Capture One.app"
	activate
	delay 1
	
	set currentDoc to current document
	
	tell currentDoc
		set targetCollection to current collection
		
		set photoListFile to choose file with prompt "Select text file with photo names:"
		set photoList to read photoListFile as «class utf8»
		set photoNames to paragraphs of photoList
		
		set albumVariants to get every variant of targetCollection
		
		set matchedVariants to {}
		set notFoundNames to {}
		
		-- Build a lookup string from the search list for fast membership check.
		-- Format: "|basename1|basename2|..." – one pass through variants instead of
		-- one full pass per search name (O(n) instead of O(n x m)).
		set searchSet to "|"
		set cleanPhotoNames to {}
		repeat with photoName in photoNames
			if photoName is not "" then
				set searchName to photoName as text
				if searchName contains "." then
					set oldDelims to AppleScript's text item delimiters
					set AppleScript's text item delimiters to "."
					set nameParts to text items of searchName
					set AppleScript's text item delimiters to oldDelims
					set searchName to ""
					repeat with i from 1 to (count of nameParts) - 1
						if i > 1 then set searchName to searchName & "."
						set searchName to searchName & (item i of nameParts)
					end repeat
				end if
				set searchSet to searchSet & searchName & "|"
				set end of cleanPhotoNames to searchName
			end if
		end repeat
		
		-- Single pass through all variants
		set foundBaseNames to "|"
		repeat with aVariant in albumVariants
			set variantName to name of aVariant
			if variantName contains "." then
				set oldDelims to AppleScript's text item delimiters
				set AppleScript's text item delimiters to "."
				set nameParts to text items of variantName
				set AppleScript's text item delimiters to oldDelims
				set baseName to ""
				repeat with i from 1 to (count of nameParts) - 1
					if i > 1 then set baseName to baseName & "."
					set baseName to baseName & (item i of nameParts)
				end repeat
			else
				set baseName to variantName
			end if
			
			if searchSet contains ("|" & baseName & "|") then
				set end of matchedVariants to aVariant
				set foundBaseNames to foundBaseNames & baseName & "|"
			end if
		end repeat
		
		-- Collect not-found names
		repeat with cleanName in cleanPhotoNames
			if foundBaseNames does not contain ("|" & cleanName & "|") then
				set end of notFoundNames to cleanName as text
			end if
		end repeat
		
		if matchedVariants is not {} then
			-- Album name
			set albumName to text returned of (display dialog "Create new album with " & (count of matchedVariants) & " found photo(s)" & return & return & "Enter album name:" default answer "New Album" buttons {"Cancel", "Create"} default button 2)
			
			-- Optional: Rating
			set ratingValue to 0
			set ratingDialog to display dialog "Set star rating for matched photos?" & return & "(1–5, or leave empty to skip)" default answer "" buttons {"Skip", "Set"} default button "Skip"
			if button returned of ratingDialog is "Set" then
				set ratingInput to text returned of ratingDialog
				if ratingInput is not "" then
					try
						set ratingValue to ratingInput as integer
						if ratingValue < 1 or ratingValue > 5 then set ratingValue to 0
					end try
				end if
			end if
			
			-- Optional: Color Tag (-1 = skip, 0 = no label, 1-6 = colors)
			set colorValue to -1
			set colorDialog to display dialog "Set color tag for matched photos?" & return & "(red / orange / yellow / green / blue / purple / no label)" & return & "Leave empty to skip." default answer "" buttons {"Skip", "Set"} default button "Skip"
			if button returned of colorDialog is "Set" then
				set colorInput to text returned of colorDialog
				if colorInput is "red" then
					set colorValue to 1
				else if colorInput is "orange" then
					set colorValue to 2
				else if colorInput is "yellow" then
					set colorValue to 3
				else if colorInput is "green" then
					set colorValue to 4
				else if colorInput is "blue" then
					set colorValue to 5
				else if colorInput is "purple" then
					set colorValue to 6
				else if colorInput is "no label" then
					set colorValue to 0
				end if
			end if
			
			-- Create album and apply
			set newAlbum to make new collection with properties {kind:album, name:albumName}
			repeat with aVariant in matchedVariants
				add inside newAlbum variants {aVariant}
				if ratingValue > 0 then
					set rating of aVariant to ratingValue
				end if
				if colorValue >= 0 then
					set color tag of aVariant to colorValue
				end if
			end repeat
			
			-- Summary
			set resultMsg to "Album '" & albumName & "' created with " & (count of matchedVariants) & " photo(s)."
			if ratingValue > 0 then set resultMsg to resultMsg & return & "Rating set: " & ratingValue & " star(s)"
			if colorValue >= 0 then set resultMsg to resultMsg & return & "Color tag set: " & colorValue
			if notFoundNames is not {} then
				set resultMsg to resultMsg & return & return & "Not found: " & (count of notFoundNames) & " photo(s)" & return
				repeat with missingName in notFoundNames
					set resultMsg to resultMsg & "• " & missingName & return
				end repeat
			end if
			
			display dialog resultMsg buttons {"OK"}
		else
			display dialog "No matching photos found in the current collection!" buttons {"OK"}
		end if
		
	end tell
	
end tell
