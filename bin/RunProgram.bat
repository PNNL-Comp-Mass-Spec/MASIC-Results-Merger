rem The following command illustrates how to:
rem  1. read a _fht.txt file for a given dataset (/I switch)
rem  2. grab the MASIC files from an alternate folder (/M switch)
rem  3. write the results on the local drive (/O:.)
rem MASICResultsMerger.exe \\proto-5\LTQ_2_DMS5\QC_Shew_08_04_pt5_18Feb09_Griffin_08-11-11\Seq200902181248_Auto366645\*_fht.txt /M:\\proto-5\LTQ_2_DMS5\QC_Shew_08_04_pt5_18Feb09_Griffin_08-11-11\SIC200902181248_Auto366644 /O:.


rem The following command illustrates how to:
rem  1. read a _fht.txt file for a given dataset (/I switch)
rem  2. grab the MASIC files from an alternate folder (/M switch)
rem  3. create separate output files for each collision mode type (/C switch)
rem  4. write the results on the local drive (/O:.)

rem MASICResultsMerger.exe \\n2.emsl.pnl.gov\dmsarch\LTQ_ETD_1_1\lowdose-b2-iTRAQ-F10-PQDETD\Seq200811050701_Auto342495\*_fht.txt /M:\\n2.emsl.pnl.gov\dmsarch\LTQ_ETD_1_1\lowdose-b2-iTRAQ-F10-PQDETD\SIC200810081144_Auto340056 /C /O:.

rem Alternatively, if all of the required files are in the same folder as MASICResultsMerger.exe, then you can use this:
rem MASICResultsMerger.exe *_fht.txt /C

rem Read lowdose_2rad_IMAC_PQD_ETD_fht.txt and create lowdose_2rad_IMAC_PQD_ETD_fht_PlusSICStats.txt using the MASIC results
MASICResultsMerger.exe lowdose_2rad_IMAC_PQD_ETD_fht.txt

rem Read lowdose_2rad_IMAC_PQD_ETD_fht.txt and create lowdose_2rad_IMAC_PQD_ETD_fht_pqd_PlusSICStats.txt using the MASIC results
MASICResultsMerger.exe lowdose_2rad_IMAC_PQD_ETD_fht.txt /C