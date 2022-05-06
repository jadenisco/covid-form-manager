# covid-form-manager

## The Covid Attestation Form Manager

This utility helps manage the covid attestation forms.

There are functions to do the following:

* split

    Splits a single pdf that contains lots of attestation forms into separate pdf files using the volunteer number in the file name.

* move

    Move the files created from split to the appropriate directories.

* Create the shortcut

    The shortcut should look something like:
    "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -File \\Cifs2\voldept$\Scripts\covid-form-manager\covid-forms.ps1"