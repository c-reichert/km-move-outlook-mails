# Move Outlook eMails into Folder via Keyboard Maestro Macro (km-move-outlook-mails)
Keyboard Maestro Macro &amp; JXA Script to move eMails into a certain folder in Outlook for Mac

This "solution" is combined of two elements:

- A KM Macro, which defines the Folder to sort into as a variable and then executes the JXA Script.
- A JXA (Javascript for Automation) Script that tells Outlook to move the selected messages to the defined folder.

## Install

- Download and move the script ([Outlook-MoveSelectedToSpecificFolder.js](Outlook-MoveSelectedToSpecificFolder.js)) to a folder you like.
- Download and install my Sample Keyboard Maestro Macro ([Sample-File-to-Folder.kmmacros](Sample-File-to-Folder.kmmacros)) and adjust the Folder name & script path as the comments in the macro instruct you. 

## Known Limitations:

- If you have mulitple Outlook accounts, then the script figures out in which account the current Message is and tries to find your destination folder in the same account. (So across account moving is currently not supported).
- Error handling is virtually non-existant in my script (... todo ...) :-)

## Credits: 

A good Question - there weren't many Outlook JXA Scripts around on the Internet - so I looked at many similar or other examples (including Applescript examples) for inspiration to get this done. Thanks to all of those who take the time to share their efforts.
