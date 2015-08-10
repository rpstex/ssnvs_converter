@echo off
setlocal enabledelayedexpansion
REM Set the filename of the SQL script that does the conversion below.
set "_converter_logic_file=ssnvs_convert_xlsx_to_txt.sql"
REM Create a variable with my preferred format of the current time to be included in the name of the output file.
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /format:list') do set datetime=%%I
set datetime=%datetime:~0,8%-%datetime:~8,6%
REM Check if powershell.exe exists within %PATH%
for %%I in (powershell.exe) do if "%%~$PATH:I" neq "" (
    set chooser=powershell "Add-Type -AssemblyName System.windows.forms|Out-Null;$f=New-Object System.Windows.Forms.OpenFileDialog;$f.InitialDirectory='%cd%';$f.Filter='All Files (*.*)|*.*';$f.showHelp=$true;$f.ShowDialog()|Out-Null;$f.FileName"
) else (
        echo Error: no powershell! omgosh!
    if not exist "!chooser!" (
        echo Error: Please install .NET 2.0 or newer, or install PowerShell.
        goto :EOF
    )
)
REM Capture the input file choice to a variable.
for /f "delims=" %%I in ('%chooser%') do set "filename=%%I"
REM make a copy of %filename% in the working directory to use for processing.
echo Converting %filename% into a flat text file compatible with SSA online batch validator.
REM planning to change lastInput to a curenttiume based variable and feed that into sqlcmd as a variable for the input file the sql logic operates on.
copy %filename% lastInput.xlsx /y
REM Execute a SQL query through SQLCMD with some switches to prevent trailing white space and status messages about affects rows from appearing.
echo Executing: sqlcmd -E -i %_converter_logic_file% -o ssnvs_output_%datetime%.txt -h-1 -W
sqlcmd -E -i %_converter_logic_file% -o ssnvs_output_%datetime%.txt -h-1 -W
echo Finished the conversion operation...
REM Debugging note: if the SQLCMD above fails, the remaining commands may fail until the dllhost.exe process is terminated.
REM Delete the copy of the file because it is not needed any longer.
del /q lastInput.xlsx
REM refactor line below..to match comments above about lastInput
echo The output file is ssnvs_output_%datetime%.txt.
echo Press any key to close this window. 
