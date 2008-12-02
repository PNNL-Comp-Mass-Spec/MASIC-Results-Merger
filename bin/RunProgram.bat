rem This command illustrates how to:
rem  1. read a _fht.txt file for a given dataset (/I switch)
rem  2. grab the MASIC files from an alternate folder (/M switch)
rem  3. create separate output files for each collision mode type (/C switch)
rem  4. write the results on the local drive (/O:.)

MASICResultsMerger.exe /I:\\proto-9\LTQ_ETD1_DMS1\lowdose-b2-iTRAQ-F10-PQDETD\Seq200811050701_Auto342495\*_fht.txt /M:\\proto-9\LTQ_ETD1_DMS1\lowdose-b2-iTRAQ-F10-PQDETD\SIC200810081144_Auto340056 /C /O:.


rem Alternatively, if all of the required files are in the same folder as MASICResultsMerger.exe, then you can use this:
rem MASICResultsMerger.exe /I:*_fht.txt /C