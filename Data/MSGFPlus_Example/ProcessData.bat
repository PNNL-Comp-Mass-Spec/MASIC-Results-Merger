@echo off

echo Example 1, process a PHRP _syn.txt file
..\..\Bin\MASICResultsMerger.exe QC_Shew_16_01-15f_13_19Nov16_Tiger_16-02-15_msgfdb_syn.txt

echo Example 2, process a .tsv file created by MzidToTsvConverter using an MSGF+ .mzid file
..\..\Bin\MASICResultsMerger.exe QC_Shew_16_01-15f_13_19Nov16_Tiger_16-02-15_msgfplus.tsv
