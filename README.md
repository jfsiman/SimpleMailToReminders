# SimpleMailToReminders
This is an AppleScript script to send a Mail message to the Reminders App

## Installation

For installation basically you:

* Create in Automator a service with no input and only for Mail.app

![Create New Workflow](/img/CreateServiceinAutomator.png?raw=true "Create Service in Automator")

* Add a step that Runs AppleScript (and replace the given code by this one)

![Add Run AppleScript](/img/RunAppleScriptCodeAutomator.png?raw=true "Add Run AppleScript")

* Save it in ~/Library/Service with some useful name like "Create Reminder from Email"
* And give it a Keyboard shortcut under System Preferences > Keyboard > Shortcuts > Services  (look in the General section)

![Create Keyboard Shortcut](/img/CreateReminderShortcut.png?raw=true "Create Keyboard Shortcut")

