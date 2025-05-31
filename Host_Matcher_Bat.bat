@echo off
setlocal

REM Configuration Parameters
set PS_SCRIPT=.\HostMatcher.ps1
set FILE_B=valid_hosts.xlsx
set JIRA_URL=http://jira.internal
set JIRA_FILTER_ID=12345
set TRIGGER_PATTERN=server names.*?encrypting is hosted on:
set OUTPUT_PREFIX=output
set EXTRACT_COLUMN=7
set ID_COLUMN=2
set REFERENCE_COLUMN=25

REM Escape quotes for PowerShell if needed
set TRIGGER_PATTERN_ESC="server names.*?encrypting is hosted on:"

REM Execute PowerShell script with parameters
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" ^
    -FileB "%FILE_B%" ^
    -JiraUrl "%JIRA_URL%" ^
    -JiraFilterId "%JIRA_FILTER_ID%" ^
    -TriggerPattern %TRIGGER_PATTERN_ESC% ^
    -OutputPrefix "%OUTPUT_PREFIX%" ^
    -ExtractColumn %EXTRACT_COLUMN% ^
    -IdColumn %ID_COLUMN% ^
    -ReferenceColumn %REFERENCE_COLUMN%

pause
