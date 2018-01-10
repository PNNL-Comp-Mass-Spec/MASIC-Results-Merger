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

echo Example 1, process a PHRP _syn.txt file
%ExePath% QC_Shew_16_01-15f_13_19Nov16_Tiger_16-02-15_msgfdb_syn.txt

echo Example 2, process a .tsv file created by MzidToTsvConverter using an MSGF+ .mzid file
%ExePath% QC_Shew_16_01-15f_13_19Nov16_Tiger_16-02-15_msgfplus.tsv

:Done
