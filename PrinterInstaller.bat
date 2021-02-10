@echo off
setlocal EnableDelayedExpansion

::TODO:
::Remote host
::Single printer add
::Interactive mode?
::Config file
::Query


set RUNDIR=%~dp0
set RUNPROMPT=%PROMPT%
set "RUNNAME=%0"
set CSVFILE=
set PRINTERNAME=
set PRINTERIP=
set PRINTERDRIVER=
set ACTION=none
set HOST=localhost
set VERBOSE=off
REM set INTERACTIVE=on
set /A SUCCESS=0
set /A FAILURE=0
set /A SKIPPED=0

prompt $G

if "%1"=="" goto :HELP

:: Retrieve arguments
:GETOPTS
	if /I "%1" == "-x" @echo on & REM Debug mode
	if /I "%1" == "-a" set ACTION=add
	if /I "%1" == "-d" set ACTION=delete
	if /I "%1" == "-v" set VERBOSE=on
	if /I "%1" == "-f" call :SETFILE "%2"
	if /I "%1" == "-h" call :SETHOST "%2"
	if /I "%1" == "-?" goto :HELP
	shift
	if not "%1" == "" goto :GETOPTS
goto :MAIN

::Describe script usage
:HELP
	echo PrinterInstaller by Bryndon Lezchuk (github.com/bryndonlezchuk)
	echo;
	echo An installation/deletion script for network printers
	echo;
	echo Usage: %RUNNAME% [-xad?] [-f csv file] [-h host]
	echo;
	echo Arguments:
	echo -x   - debug mode
	echo -a   - add printer(s)
	echo -d   - delete printer(s)
	echo -f   - csv file with preset printer info
	echo -h   - host to connect to. Default is localhost
	echo -?   - display command usage
	echo;
	echo Examples:
	echo %RUNNAME% -a -f "file.csv"
	echo %RUNNAME% -d -f "file.csv"
	echo %RUNNAME% -a -f "file.csv" -h "\\hostname"
	echo;
	echo Remarks:
	echo Expected format of the CSV file is PrinterName,PrinterIP,PrinterDriver
goto :EOF

:: Main body of script goes here
:MAIN
::if %ACTION%==none call :EXITERROR "Expecting add (-a) or delete (-d) argument"

::Check for any errors in arguments
	if not errorlevel 0 goto :CLEANUP
	echo Starting script...

	for /f "tokens=1-4 delims=," %%A IN (%CSVFILE%) do (
		set "PRINTERNAME=%%A"
		set "PRINTERIP=%%B"
		set "PRINTERDRIVER=%%C"
		
		if %ACTION%==add call :ADD "%%A", "%%B", "%%C"
		if %ACTION%==delete call :DELETE "%%A", "%%B"
		
		if not errorlevel 0 goto :CLEANUP
	)

:CLEANUP
	echo;
	if %ACTION%==add (
		echo !SUCCESS! successfull installations
		echo !FAILURE! failed installations
		echo !SKIPPED! skipped installations
	) else if %ACTION%==delete (
		echo %SUCCESS% successfull deletions
		echo %FAILURE% failed deletions
		echo %SKIPPED% skipped deletions
	)
	echo;
	echo End of script

	cd %RUNDIR%
	prompt %RUNPROMPT%
goto :EOF

::Add a printer to local machine
:: 	Input 1: Printer Name
:: 	Input 2: Printer IP
:: 	Input 3: Printer Driver
:ADD
	echo;
	echo ----------------------------------------
	echo Adding printer "%~1" at IP %~2 with driver "%~3"
	
	cd %WINDIR%\System32\Printing_Admin_Scripts\en-US\
	
	::Check for driver
	set DRVRRESULT=
	for /f "tokens=* USEBACKQ" %%F in (`cscript prndrvr.vbs -l ^| find "%~3"`) do set DRVRRESULT=%%F
	if "%DRVRRESULT%"=="" (
		::Driver not found
		echo Driver "%~3" not found, skipping printer install
		echo ----------------------------------------
		call :SKIP
		exit /b 0
	) else (
		::Driver is installed and found
		echo Driver "%~3" is installed
	)
	
	::Check for port
	set PORTRESULT=
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnport.vbs -l ^| find "Port name %~2"`) do set PORTRESULT=%%F
	if "%PORTRESULT%"=="" (
	
		REM Port not found
		REM Create the port
	
		echo Port not found for "%~2", adding port
		cscript prnport.vbs -a -r %~2 -h %~2 -o raw -n 9100 | find "Created"
		
		if not errorlevel 0 (
			call :FAIL
			exit /b %ERRORLEVEL%
		)
	) else echo Port "%~2" already created, skipping
	
	::Install printer
	echo Installing "%~1"
	cscript prnmngr.vbs -a -p "%~1" -m "%~3" -r "%~2" | find "Added"
	
	if not errorlevel 0 (call :FAIL) else (call :PASS)
	
	echo ----------------------------------------
::	endlocal
exit /b %ERRORLEVEL%

::Delete a printer from local machine
:: 	Input 1: Printer Name
:: 	Input 2: Printer IP
:DELETE
::	setlocal
	echo;
	echo ----------------------------------------
	
	cd %WINDIR%\System32\Printing_Admin_Scripts\en-US\
	
	::Remove printer
	set PRNRESULT=
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnmngr.vbs -l ^| find "%~1"`) do set PRNRESULT=%%F
	if "%PRNRESULT%"=="" (
		::Printer not present
		echo Printer "%~1" not found, checking port...
	) else (
		::Printer is present
		echo Removing printer "%~1"...
		cscript prnmngr.vbs -d -p "%~1" | find "Deleted"
		
		if not errorlevel 0 (
			call :FAIL
			exit /b %ERRORLEVEL%
		)
	)
	
	::Remove port
	set PORTRESULT=
	for /f "tokens=* USEBACKQ" %%F in (`cscript prnport.vbs -l ^| find "Port name %~2"`) do set PORTRESULT=%%F
	if "%PORTRESULT%"=="" (
		::Port not present
		echo Port "%~2" not found, skipping
		call :SKIP
	) else (
		::Port is present
		echo Removing port "%~2"...
		cscript prnport.vbs -d -r "%~2" | find "Deleted"
		
		if not errorlevel 0 (call :FAIL) else (call :PASS)
	)
	
	echo ----------------------------------------
::	endlocal
exit /b %ERRORLEVEL%

::Set the host machine
:: Input 1: Host to be set
:SETHOST
	set "HOST=%~1""
exit /b %ERRORLEVEL%

::Set the filepath for the CSV file
:: 	Input 1: CSV file to use
:SETFILE
	call :CHECKFILE "%~1"
	set "CSVFILE=%~1"
exit /b %ERRORLEVEL%

::Verify file existance
:: 	Input 1: File to validate
:CHECKFILE
	if not exist "%~1" call :EXITERROR "No such file"
exit /b %ERRORLEVEL%

::Increment the success counter
:PASS
	set /A SUCCESS+=1
exit /b 0

::Increment the failure counter
:FAIL
	set /A FAILURE+=1
exit /b 0

::Increment the skipped counter
:SKIP
	set /A SKIPPED+=1
exit /b 0

:VECHO
	if %VERBOSE%==on echo %~1
exit /b 0

::Displays an error message and exits script (easier said than done)
:: 	Input 1: Message
:EXITERROR
	echo ERROR: %~1
exit /b 1