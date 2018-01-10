@echo off

set ExePath=MASICResultsMerger.exe

if exist %ExePath% goto DoWork
if exist ..\%ExePath% set ExePath=..\%ExePath% && goto DoWork
if exist ..\..\Bin\%ExePath% set ExePath=..\..\Bin\%ExePath% && goto DoWork

echo Executable not found: %ExePath%
goto Done

:DoWork
echo.
echo Procesing with %ExePath%
echo.

echo Process a Mage Extractor file
%ExePath% AThaliana_Sequest_Syn.txt /Mage

:Done
