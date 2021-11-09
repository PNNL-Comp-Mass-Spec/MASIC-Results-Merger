# MASIC Results Merger

This program merges the contents of a tab-delimited peptide hit results file from PHRP
(for MS-GF+, MaxQuant, X!Tandem, etc.) with the corresponding MASIC results files,
appending the relevant MASIC stats to each peptide hit result,
writing the merged data to a new tab-delimited text file.

It also supports TSV files, e.g. as created by the 
[MzidToTsvConverter](https://github.com/PNNL-Comp-Mass-Spec/Mzid-To-Tsv-Converter)

If the input directory includes a MASIC _ReporterIons.txt file, 
the reporter ion intensities will also be included in the new text file.

## Console Switches

MASICResultsMerger is a console application, and must be run from the Windows command prompt.

```
MASICResultsMerger.exe 
  InputFilePathSpec 
  [/M:MASICResultsDirectoryPath] [/O:OutputDirectoryPath]
  [/N:ScanNumberColumn] [/C] [/Mage] [/Append]
  [/DartID]
  [/S:[MaxLevel]] [/A:AlternateOutputDirectoryPath] [/R]
  [/Trace]
```

The input file should be a tab-delimited file where one column has scan numbers.
By default, this program assumes the second column has scan number, but the `/N`
switch can be used to change this (see below).

Common input files are
* Peptide Hit Results Processor (https://github.com/PNNL-Comp-Mass-Spec/PHRP) tab-delimited files
  * MS-GF+ syn/fht file (`_msgfplus_syn.txt` or `_msgfplus_fht.txt`)
  * SEQUEST Synopsis or First-Hits file (`_syn.txt `or `_fht.txt`)
  * X!Tandem `_xt.txt` file
* MzidToTSVConverter (https://github.com/PNNL-Comp-Mass-Spec/Mzid-To-Tsv-Converter) .TSV files
  * This is a tab-delimited text file created from a `.mzid` file (e.g. from MS-GF+)

If the MASIC result files are not in the same directory as the input file, use `/M` 
to define the path to the correct directory. 

The output directory switch is optional.  If omitted, the output file will be 
created in the same directory as the input file. 

Use `/N` to change the column number that contains scan number in the input file.
The default is 2 (meaning `/N:2`).

When reading data with _ReporterIons.txt files, you can use `/C` to specify that a 
separate output file be created for each collision mode type in the input file 
(typically pqd, cid, and etd).

Use `/Mage` to specify that the input file is a results file from Mage Extractor.
This file will contain results from several analysis jobs; the first column in 
this file must be Job and the remaining columns must be the standard Synopsis or 
First-Hits columns supported by PHRPReader.  In addition, the input directory must
have a file named InputFile_metadata.txt (this file will have been auto-created 
by Mage Extractor).

Use `/Append` to merge results from multiple datasets together as a single file; 
this is only applicable when the InputFilePathSpec includes a * wildcard and 
multiple files are matched. The merged results file will have DatasetID values of 
1, 2, 3, etc. along with a second file mapping DatasetID to Dataset Name.

Use `/DartID` to only list each peptide once per scan. The Protein column will list
the first protein, while the Proteins column will be a comma separated list of
all of the proteins. This format is compatible with DART-ID
(https://pubmed.ncbi.nlm.nih.gov/31260443/)

Use `/S` to process all valid files in the input directory and subdirectories. 
Include a number after `/S` (like `/S:2`) to limit the level of subdirectories to examine.

When using `/S`, you can redirect the output of the results using `/A` to specify an alternate output directory.

When using `/S`, you can use `/R` to re-create the input directory hierarchy in the alternate output directory (if defined).

Use `/Trace` to show additional debug messages while searching for input files

## Contacts

Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) \
E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov\
Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics/
Source code: https://github.com/PNNL-Comp-Mass-Spec/MASIC-Results-Merger

## License

MASICResultsMerger is licensed under the Apache License, Version 2.0; you may not use this 
file except in compliance with the License.  You may obtain a copy of the 
License at https://opensource.org/licenses/Apache-2.0
