@echo off

set ExePath=MASICResultsMerger.exe

if exist %ExePath% goto DoWork
if exist ..\%ExePath% set ExePath=..\%ExePath% && goto DoWork
if exist ..\..\Bin\%ExePath% set ExePath=..\..\Bin\%ExePath% && goto DoWork

echo Executable not found: %ExePath%
goto Done

:DoWork
echo.
echo Processing with %ExePath%
echo.

echo Process a PHRP _fht.txt file
%ExePath% QC_Shew_Excerpt_fht.txt

:Done
