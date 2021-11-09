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

echo Example 1, process a PHRP _fht.txt file
%ExePath% QC_pilot_04_25Apr16_msgfdb_fht.txt

echo Example 2, process a .tsv file created by MzidToTsvConverter using an MSGF+ .mzid file
%ExePath% QC_pilot_04_25Apr16_msgfplus.tsv

:Done
