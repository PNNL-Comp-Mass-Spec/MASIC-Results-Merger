@echo off

echo Example 1, process a _syn.txt file created by PeptideHitResultsProcessor (PHRP) using a MSGF+ results
..\..\Bin\MASICResultsMerger.exe Elite_QCmam_test9_msgfdb_syn.txt

echo Example 2, create separate output files for each collision mode
..\..\Bin\MASICResultsMerger.exe Elite_QCmam_test9_msgfdb_syn.txt /C

rem echo Example 3, create separate output files for each collision mode, use the .TSV file from MzidToTsvConverter
rem ..\..\Bin\MASICResultsMerger.exe Elite_QCmam_test9_msgfplus.tsv /C