


REM Login 
REM epmautomate importBalances "EBS-ALL-Snapshot" "March 2021"

REM SETLOCAL ENABLEDELAYEDEXPANSION

FOR /f "tokens=2-4 delims=/ " %%a IN ('date /t') DO (
	set MONTH=%%a
	set YEAR=%%c
	)

ECHO %MONTH% - %YEAR%

if %MONTH%==01 set PERIOD=January %YEAR%
if %MONTH%==02 set PERIOD=Febuary %YEAR%
if %MONTH%==03 set PERIOD=March %YEAR%
if %MONTH%==04 set PERIOD=April %YEAR%
if %MONTH%==05 set PERIOD=May %YEAR%
if %MONTH%==06 set PERIOD=June %YEAR%
if %MONTH%==07 set PERIOD=July %YEAR%
if %MONTH%==08 set PERIOD=August %YEAR%
if %MONTH%==09 set PERIOD=September %YEAR%
if %MONTH%==10 set PERIOD=October %YEAR%
if %MONTH%==11 set PERIOD=November %YEAR%
if %MONTH%==12 set PERIOD=December %YEAR%

echo %PERIOD%