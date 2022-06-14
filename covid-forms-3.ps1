Write-Host "The Covid Form Manager"
//cifs2/voldept$/Scripts/covid-form-manager/.venv3/Scripts/Activate.ps1
Set-Location //cifs2/voldept$/Scripts/covid-form-manager
python.exe ./covid-forms.py -d split
# python.exe ./covid-forms.py split '.\forms\Covid Forms 04-19-2022.pdf'
# deactivate
# Write-Host -NoNewLine 'Press any key to exit...';
# $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

<#
(.venv) PS Microsoft.PowerShell.Core\FileSystem::\\cifs2\voldept$\Scripts\covid-form-manager>  & '\\cifs2\voldept$\Scripts\covid-form-manager\.venv\Scripts\python.exe' 'c:\Users\JD1060\.vscode\extensions\ms-python.python-2022.4.1\pythonFiles\lib\python\debugpy\launcher' '53356' '--' '\\cifs2\voldept$\Scripts\covid-form-manager\covid-forms.py' '-d' 'split' './forms/Covid-Forms-12-22-2021.pdf'
DEBUG:root:mgh_util(Namespace(debug=1, create=None, filename='./forms/Covid-Forms-12-22-2021.pdf', func=<function split at 0x000001B563AAE3B0>))
#>