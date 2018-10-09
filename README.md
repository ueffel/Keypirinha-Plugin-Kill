Keypirinha Plugin Kill
=======================

This is a package that extends the fast keystroke launcher keypirinha (http://keypirinha.com/) with
a command to kill a running processes.

## Usage

Type the trigger "kill" in the launch box and you'll see the item `Kill:`. After hitting Enter or
Tab, you are present with a list of running processes where you can choose a process to kill.

There some alternative actions, if you hit Ctrl+Enter which should be self-explaining.

## Installation

### With [PackageControl](https://github.com/ueffel/Keypirinha-PackageControl)

Install Package "Keypirinha-Plugin-Kill"

### Manually

* Download the `Kill.keypirinha-package` from the [releases](https://github.com/ueffel/Keypirinha-Plugin-Kill/releases/latest).
* Copy the file into `%APPDATA%\Keypirinha\InstalledPackages` (installed mode) or
  `<Keypirinha_Home>\portable\Profile\InstalledPackages` (portable mode)

## Acknowledgements

Parts of the code are taken from [Keypirinha's Packages Repository](https://github.com/Keypirinha/Packages).
Specifically the alttab.py library file to obtain visible windows is taken from
[here](https://github.com/Keypirinha/Packages/blob/9e1a0645b16577a8cefd64510cbc15690ae8ceeb/TaskSwitcher/lib/alttab.py)
