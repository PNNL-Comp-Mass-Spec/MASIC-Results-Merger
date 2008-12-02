The MASIC Results Merger reads the contents of a tab-delimited peptide hit 
results file (e.g. from Sequest, XTandem, or Inspect) and merges that 
information with the corresponding MASIC results files, appending the 
relevant MASIC stats for each peptide hit result

The input file will typically be a first hits or synopsis file, where the
second column is scan number.  However, any tab-delimited text file can be
used for the input; you can define which column contains the scan number using
the /N switch, e.g. /N:1 indicates the first column contains scan number.

The program will use the name of the input file to find the corresponding
MASIC result files, looking in the same folder as the input file.  To specify
a different folder, use the /M switch.  Two MASIC files are required: the
_ScanStats.txt file and the _SICStats.txt file.  If the _ReporterIons.txt file
is also present, then it will also be read.  

When the _ReporterIons.txt file is present, you can use the /C switch to 
instruct the program to create separate output files for each collision mode
type; this would be useful if the dataset had a mix of collision mode types,
for example pqd and etd.

-------------------------------------------------------------------------------
Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
Copyright 2008, Battelle Memorial Institute.  All Rights Reserved.

E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
Website: http://ncrr.pnl.gov/ or http://omics.pnl.gov
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

All publications that result from the use of this software should include 
the following acknowledgment statement:
 Portions of this research were supported by the W.R. Wiley Environmental 
 Molecular Science Laboratory, a national scientific user facility sponsored 
 by the U.S. Department of Energy's Office of Biological and Environmental 
 Research and located at PNNL.  PNNL is operated by Battelle Memorial Institute 
 for the U.S. Department of Energy under contract DE-AC05-76RL0 1830.

Notice: This computer software was prepared by Battelle Memorial Institute, 
hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
Department of Energy (DOE).  All rights in the computer software are reserved 
by DOE on behalf of the United States Government and the Contractor as 
provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
SOFTWARE.  This notice including this sentence must appear on any copies of 
this computer software.
