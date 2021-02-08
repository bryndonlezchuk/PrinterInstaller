@echo off
prompt $G
title Printer Installer

set CSVFILE=
set PRINTERNAME=
set PRINTERIP=
set PRINTERDRIVER=
set ACTION=none
set RUNDIR=%~dp0
set HOST=localhost

:: Retrieve arguments
:GETOPTS
	if /I "%1" == "-v" @echo on
	if /I "%1" == "-a" set ACTION=add
	if /I "%1" == "-d" set ACTION=delete
	if /I "%1" == "-f" call :SETFILE "%2"
	if /I "%1" == "-h" call :SETHOST "%2"
shift
if not "%1" == "" goto :GETOPTS

:: Main body of script goes here
:MAIN
if %ACTION%==none call :EXITERROR "Expecting add (-a) or delete (-d) argument"
if not %ERRORLEVEL%==0 goto :EOF
for /f "tokens=1-3 delims=," %%A IN (%CSVFILE%) do (
	if %ACTION%==add call :ADD "%%A", "%%B", "%%C"
	if %ACTION%==delete call :DELETE "%%A", "%%B"
	
	if not %ERRORLEVEL%==0 goto :EOF
)

cd %RUNDIR%
echo End of script
goto :EOF

::Add a printer to local machine
:: 	Input 1: Printer Name
:: 	Input 2: Printer IP
:: 	Input 3: Printer Driver
:ADD
	setlocal
	echo ""
	echo ""
	echo Adding printer "%~1" at IP %~2 with driver "%~3"
	
	cd %WINDIR%\System32\Printing_Admin_Scripts\en-US\
	
	::Check for driver
	for /f "tokens=* USEBACKQ" %%F in (`cscript prndrvr.vbs -l ^| find "%~3"`) do set DRVRRESULT=%%F
	if "%DRVRRESULT%"=="" (
		echo Driver "%~3" not found, skipping printer install
		exit /b 0
	) else (
		echo Driver "%~3" is installed
	)
	
	::Check for port
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnport.vbs -l ^| find "Port name %~2"`) do set PORTRESULT=%%F
	if "%PORTRESULT%"=="" (
		echo Port not found for "%~2", adding port
		cscript prnport.vbs -a -r %~2 -h %~2 -o raw -n 9100
	) else echo Port "%~2" already installed, skipping
	
	::Install printer
	echo Installing "%~1"
	cscript prnmngr.vbs -a -p "%~1" -m "%~3" -r "%~2"
	
	endlocal
exit /b 0

::Delete a printer from local machine
:: 	Input 1: Printer Name
:: 	Input 2: Printer IP
:DELETE
	setlocal
	echo ""
	echo ""
	
	cd %WINDIR%\System32\Printing_Admin_Scripts\en-US\
	
	::Remove printer
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnmngr.vbs -l ^| find "%~1"`) do set PRNRESULT=%%F
	if "%PRNRESULT%"=="" (
		::Printer not present
		echo Printer "%~1" not found, checking port...
	) else (
		::Printer is present
		echo Removing printer "%~1"
		cscript prnmngr.vbs -d -p "%~1"
	)
	
	::Remove port
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnport.vbs -l ^| find "Port name %~2"`) do set PORTRESULT=%%F
	if "%PORTRESULT%"=="" (
		::Port not present
		echo Port "%~2" not found, skipping
	) else (
		::Port is present
		echo Removing port "%~2"
		cscript prnport.vbs -d -r "%~2"
	)
exit /b %ERRORLEVEL%

::Set the host machine
:: Input 1: Host to be set
:SETHOST
	echo Feature not yet implemented.
exit /b %ERRORLEVEL%

::Set the filepath for the CSV file
:: 	Input 1: CSV file to use
:SETFILE
	call :CHECKFILE %~1
	set "CSVFILE=%~1
exit /b %ERRORLEVEL%

::Verify file existance
:: 	Input 1: File to validate
:CHECKFILE
	if not exist %~1 call :EXITERROR "No such file"
exit /b %ERRORLEVEL%

::Displays an error message and exits script (easier said than done)
:: 	Input 1: Message
:EXITERROR
	echo ERROR: %~1
exit /b 1