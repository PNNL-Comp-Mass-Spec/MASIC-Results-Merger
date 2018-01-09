@echo off

echo Example 1, process a PHRP _fht.txt file
..\..\Bin\MASICResultsMerger.exe QC_pilot_04_25Apr16_msgfdb_fht.txt

echo Example 2, process a .tsv file created by MzidToTsvConverter using an MSGF+ .mzid file
..\..\Bin\MASICResultsMerger.exe QC_pilot_04_25Apr16_msgfplus.tsv
