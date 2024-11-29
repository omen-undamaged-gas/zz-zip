@echo off
rem white-black
color f0
set sendto="%AppData%\Microsoft\Windows\SendTo"
set exe7z="%ProgramFiles%\7-Zip\7z.exe"

echo Copy zz-zip files to %sendto%
@pushd %~dp0
copy zz-zip* %sendto% /-Y
@popd

echo.
echo Check status of 7-Zip installation...
if exist %exe7z% (
	rem white-green
	color f2
	echo %exe7z% exists. GOOD JOB^!
) else (
	rem white-red
	color f4
	echo %exe7z% does NOT exist.
	echo Please get 7-Zip from https://www.7-zip.org/
)
echo.
pause
