Keypirinha Plugin Kill
=======================

This is a package that extends the semantic keystroke launcher keypirinha (http://keypirinha.com/) with a command to kill a running processes.

## Usage

Type the trigger "kill" in the launch box and you'll see the item "Kill:". After hitting Enter or Tab, you are present with a list of running processes where you can choose a process to kill.
After hitting Enter the selected process is killed by running "taskkill /F /IM &lt;selected item process name&gt;"
For example: "taskkill /F /IM notepad.exe"

There some alternative actions, if you hit Ctrl+Enter which should be self-explaining.

## Installation
See Release Notes.
