@echo off
echo ---- Manage Covid Forms ----
echo:
set VROOT=z:\Developer\Windows\testroot
set PROOT=scripts\covid-form-manager
set VENV=%VROOT%\%PROOT%\.venv
%VENV%\Scripts\activate.bat & python %VROOT%\%PROOT%\covid-forms.py -h & timeout /t 30 & deactivate
exit
