rem The following command illustrates how to read a _fht.txt file for a given dataset
rem This assumes the MASIC result files (_SICstats.txt and _ScanStats.txt) are in the same folder; use /M if they're in an alternate folder
MASICResultsMerger.exe /I:QC_Shew_Excerpt_fht.txt

rem Alternatively, if all of the required files are in the same folder as MASICResultsMerger.exe, then you can use this:
rem MASICResultsMerger.exe /I:*_fht.txt

rem  You can optionally use /C to create separate output files for each collision mode type
