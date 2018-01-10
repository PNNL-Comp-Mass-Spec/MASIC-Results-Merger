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

echo Example 1, process a _syn.txt file created by PeptideHitResultsProcessor (PHRP) using a MSGF+ results
%ExePath% Elite_QCmam_test9_msgfdb_syn.txt

echo Example 2, create separate output files for each collision mode
%ExePath% Elite_QCmam_test9_msgfdb_syn.txt /C

rem echo Example 3, create separate output files for each collision mode, use the .TSV file from MzidToTsvConverter
rem %ExePath% Elite_QCmam_test9_msgfplus.tsv /C

:Done
