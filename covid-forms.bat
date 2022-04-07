@echo off
echo ---- Manage Covid Forms ----
echo:
cd forms
set /P filename=Please enter the file to split: 
cd ..
set VROOT=z:\Developer\Windows\testroot
set PROOT=scripts\covid-form-manager
set VENV=%VROOT%\%PROOT%\.venv
%VENV%\Scripts\activate.bat & python %VROOT%\%PROOT%\covid-forms.py split %VROOT%\%PROOT%\forms\%filename% & timeout /t 30 & deactivate
timeout /t 30
exit
