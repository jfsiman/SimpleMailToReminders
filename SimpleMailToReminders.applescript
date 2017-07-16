-------------------------------------------------------
----------         Mail To Reminders         ----------
-------------------------------------------------------
-- Author: José Simán
-- Using code from: James Gibbard (http://jgibbard.me.uk)
-- Website: 
-- Version: 0.1

-- Last Updated: 15/07/2017

-- Description:
# An AppleScript to create tasks (in Reminders) from
# emails in Mail. It can be used as a script 
# called from a shortcut key.

-- Dependent Apps:
# - Mail (http://www.apple.com)
# - Reminders v4.0+ 

-------------------------------------------------------
--           PROPERTIES TO BE ADJUSTED             --
-------------------------------------------------------

-- Property: defaultTag
# Sets the default tag that is assigned to
# every task that is created from a note
property defaultTag : "Email"

-- Property: userTags
# Allows the ability to specify if user defined
# tags should be turned ON (1) or OFF (0)
property userTags : 1

-- Property: tagList
# Sets out the userTags that can be selected
# when creating a task from a message 
property tagList : {"Needs Reply", "Needs Action", "Needs Review", "Follow Up", "Create Tasks"}

-- Note: The flag colour is determined by the index 
# position of the userTags in the tagList.
# Example: 'Needs Reply' is index 1 = Orange Flag
# NB: The Reg flag is not used, as it is 
# the default flag for flagged emails.

-- Property: setFlags
# Allows the ability to set flags in Mail.
# It can be turned ON (1) or OFF (0).
property setFlags : 1

-- Property: showNotification
# Sets if notifications are shown when
# a task is created. ON (1) or OFF (0)
property showNotification : 1

-- Set this before running the Script
# Set the name of the default reminder list (depends on your OS Language)
set DefaultReminderList to "Work"

-- I usually like to be reminded early in the morning.
# Set the default reminder time in hours after midnight, I suggest any number between 0,5 and 23,5
# for a reminder at "8:00 am" set "8", for "3 PM" or "15:00" set "15", for "8h45" set "8,75"
set defaultReminderTime to "9"

-- I usually like to follow up on reminders the following day
# Set the default reminder date
# these are the possible choices: "Tomorrow", "2 Days", "3 Days", "4 Days", "End of Week", "Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Specify"
set defaultReminder to "Tomorrow"


-------------------------------------------------------
--  GLOBAL VARIABLES (NOT TO BE AJUSTED)  --
-------------------------------------------------------

-- Global Variable: theCustomTags
# Sets the var for holding all the custom
# tags to be blank on initialisation
set theCustomTags to ""

-- Global Variable: theMessageURL
# Sets the var for holding the message URL
# be 'missing value' on initialisation
set theMessageURL to missing value

-------------------------------------------------------
--  MAIN SCRIPT (EXECUTION STARTS HERE)   --
-------------------------------------------------------

-- Connect to Mail 
tell application "Mail"
	try
		-- Get the selected note in Mail
		set selectedMessages to selection
		
		-- Get the Subject from the selected message
		set theMessageSubject to (subject of item 1 of selectedMessages) as string
		
		-- Get the MessageID from the selected message
		set theMessageID to (message id of item 1 of selectedMessages) as string

		-- Get the Sender from the selected message
		set theMessageSender to (sender of item 1 of selectedMessages) as string
		
	on error errMsg
		-- Throw error and exit, if we can't select the note in Evernote
		-- Show Notification in OS X
		if (showNotification = 1) then
			display notification "ERROR: Couldn't create task!"
			delay 1
		else
			display dialog "ERROR: Couldn't create task!"
		end if
		
		return
	end try
	
	-- if the property 'userTags' is set to 1
	if (userTags = 1) then
		-- Ask the user which tags they would like to apply to the task
		set theTags to (choose from list tagList with prompt "Select the tags to include:" with multiple selections allowed)
		
		-- If any tags are seleceted the user, add them to the string of tags
		if theTags is not false then
			repeat with theTag in theTags
				set theCustomTags to theCustomTags & "," & theTag
				
				-- If setFlags is ON, then set the falg in Mail
				if (setFlags = 1) then
					-- Get the tag list posision to work out the flag index
					set flagIndex to my list_position(theTag as string, tagList)
					-- Set the flag index in Mail
					set flag index of (item 1 of selectedMessages) to flagIndex
				end if
			end repeat
		else
			error number -128
		end if
	end if
end tell

-- Format the Message URL with the MessageID
set theMessageURL to "message://%3c" & my replace_chars(theMessageID, "%", "%25") & "%3e"

-- Format the Task Notes(add linefeeds) Needs (this is not needed in Reminders)
-- set theTaskNotes to theMessageURL & return & linefeed
set theTaskNotes to theMessageURL 

-- Remove any &'s from the subject
set the theTask to my replace_chars(theMessageSubject, "&", "and")

set first_name to ""
set last_name to ""
set senderName to ""

-- Get the senders full name
if (existsInString(",", theMessageSender)) then
	set AppleScript's text item delimiters to ","
	try
		copy every text item of theMessageSender to {last_name, first_name}
	end try
	
	set AppleScript's text item delimiters to "\""
	try
		copy every text item of first_name to {first_name, foo}
	end try
	set AppleScript's text item delimiters to ""
	set the last_name to my replace_chars(last_name, "\"", "")
	
	set senderName to first_name & " " & last_name
else
	set AppleScript's text item delimiters to " <"
	try
		copy every text item of theMessageSender to {senderName, foo}
	end try
	set AppleScript's text item delimiters to ""
end if

-- Set variables based on default values defined above
# for all the other options, calculate the date based on the current date
set reminderDate to defaultReminder
set remindMeDate to my chooseRemindMeDate(reminderDate)
set time of remindMeDate to 60 * 60 * defaultReminderTime

#default list name in Reminders
set RemindersList to DefaultReminderList


-- Create the tasks in 2Do (using the app's URL)
-- do shell script "open 'twodo://x-callback-url/add?" & "task=" & theTask & "&forlist=" & "&note=" & theTaskNotes & "&tags=" & defaultTag & theCustomTags & "," & senderName & "'"

tell application "Reminders"
	
	tell list RemindersList
		# create new reminder with proper due date, subject name and the URL linking to the email in Mail
		make new reminder with properties {name:theTask, remind me date:remindMeDate, body:theTaskNotes}
		
	end tell
	
end tell

-- Show Notification in OS X
if (showNotification = 1) then
	display notification theTask with title "Task added to Reminders Successfully!"
	delay 1
end if

-------------------------------------------------------
--             FUNCTION: Replace Chars               --
-------------------------------------------------------
on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars

-------------------------------------------------------
--               FUNCTION: List Position             --
-------------------------------------------------------
on list_position(this_item, this_list)
	repeat with i from 1 to the count of this_list
		if item i of this_list is this_item then return i
	end repeat
	return 0
end list_position

-------------------------------------------------------
--             FUNCTION: Exists in String            --
-------------------------------------------------------
on existsInString(subString, thisString)
	set charExists to false
	set i to offset of subString in thisString
	if (i > 0) then
		set charExists to true
	end if
	return charExists
end existsInString

-------------------------------------------------------
--             FUNCTION: chooseRemindMeDate          --
-------------------------------------------------------

# date calculation with the selection from the dialogue if there is one
# use to set the initial and the re-scheduled date
on chooseRemindMeDate(selectedDate)
	if selectedDate = "Tomorrow" then
		# add 1 day and set time to 9h into the day = 9am
		set remindMeDate to (current date) + 1 * days
		
	else if selectedDate = "2 Days" then
		set remindMeDate to (current date) + 2 * days
		
	else if selectedDate = "3 Days" then
		set remindMeDate to (current date) + 3 * days
		
	else if selectedDate = "4 Days" then
		set remindMeDate to (current date) + 4 * days
		
	else if selectedDate = "End of Week" then
		# end of week means Thursday in terms of reminders
		# get the current day of the week
		set curWeekDay to weekday of (current date) as string
		if curWeekDay = "Monday" then
			set remindMeDate to (current date) + 3 * days
		else if curWeekDay = "Tuesday" then
			set remindMeDate to (current date) + 2 * days
		else if curWeekDay = "Wednesday" then
			set remindMeDate to (current date) + 1 * days
			# if it's Thursday, I'll set the reminder for Friday
		else if curWeekDay = "Thursday" then
			set remindMeDate to (current date) + 1 * days
			# if it's Friday I'll set the reminder for Thursday next week
		else if curWeekDay = "Friday" then
			set remindMeDate to (current date) + 6 * days
		else if curWeekDay = "Saturday" then
			set remindMeDate to (current date) + 5 * days
		else if curWeekDay = "Sunday" then
			set remindMeDate to (current date) + 4 * days
		end if
		
	else if selectedDate = "Next Monday" then
		set curWeekDay to weekday of (current date) as string
		if curWeekDay = "Monday" then
			set remindMeDate to (current date) + 7 * days
		else if curWeekDay = "Tuesday" then
			set remindMeDate to (current date) + 6 * days
		else if curWeekDay = "Wednesday" then
			set remindMeDate to (current date) + 5 * days
		else if curWeekDay = "Thursday" then
			set remindMeDate to (current date) + 4 * days
		else if curWeekDay = "Friday" then
			set remindMeDate to (current date) + 3 * days
		else if curWeekDay = "Saturday" then
			set remindMeDate to (current date) + 2 * days
		else if curWeekDay = "Sunday" then
			set remindMeDate to (current date) + 1 * days
		end if
		
	else if selectedDate = "1 Week" then
		set remindMeDate to (current date) + 7 * days
		
	else if selectedDate = "2 Weeks" then
		set remindMeDate to (current date) + 14 * days
		
	else if selectedDate = "1 Month" then
		set remindMeDate to (current date) + 28 * days
		
	else if selectedDate = "2 Months" then
		set remindMeDate to (current date) + 56 * days
		
	else if selectedDate = "3 Months" then
		set remindMeDate to (current date) + 84 * days
		
	else if selectedDate = "Specify" then
		# adapt the date format suggested with what is configured in the user's 'Language/Region'-Preferences
		set theDateSuggestion to (short date string of (current date))
		set theDateInput to text returned of (display dialog "Type the date for the reminder (e.g. '" & theDateSuggestion & "'):" default answer theDateSuggestion buttons {"Cancel", "OK"} default button "OK")
		try
			set remindMeDate to date theDateInput
		on error
			set remindMeDate to (current date) + 1 * days
			(display dialog "There was an error with the date input provided: '" & theDateInput & "'. The reminder was set to tomorrow." with title "Error: '" & theDateInput & "'")
		end try
	end if
	
	return remindMeDate
end chooseRemindMeDate
