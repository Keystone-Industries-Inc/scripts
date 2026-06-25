REM Replace {Instance Name}, {Database Name}, and {Backup Directory Path} with your actual values before running the script.
@echo off
:: Set environment variables
set INSTANCE=.\{Instance Name}}
set DBNAME={Database Name}
set BACKUP_DIR={Backup Directory Path}

:: Generate date stamp (YYYYMMDD) for unique file names
set DATESTAMP=%date:~10,4%%date:~4,2%%date:~7,2%
set FILEPATH=%BACKUP_DIR%\%DBNAME%_%DATESTAMP%.bak

:: Ensure backup directory exists
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"

:: Execute the sqlcmd backup
sqlcmd -S %INSTANCE% -E -Q "BACKUP DATABASE [%DBNAME%] TO DISK='%FILEPATH%' WITH INIT, FORMAT, STATS=10"
