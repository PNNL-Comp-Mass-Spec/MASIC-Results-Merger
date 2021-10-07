# __<span style="color:#D57500">MASIC Results Merger</span>__
Reads the contents of a tab-delimited peptide hit results file (e.g. from Sequest, XTandem, Inspect, or MSGF+) and merges that information with the corresponding MASIC results files, appending the relevant MASIC stats for each peptide hit result.

### Description
The input file will typically be a Synopsis or First Hits file (created using the [Peptide Hit Results Processor](https://pnnl-comp-mass-spec.github.io/PHRP/)), where the second column is scan number. However, any tab-delimited text file can be used for the input; you can define which column contains the scan number using the /N switch, e.g. /N:1 indicates the first column contains scan number.

If MASIC produced a _ReporterIons.txt file (created for ITRAQ data), then you can use the /C switch with the MASIC Results Merger to instruct the program to create separate output files for each collision mode type; this would be useful if the dataset had a mix of collision mode types, for example PQD and ETD.

When the InputFilePathSpec includes a * wildcard and multiple files are matched, you can use /Append to merge results from multiple datasets together as a single file.

### Downloads
* [Latest version](https://github.com/PNNL-Comp-Mass-Spec/MASIC-Results-Merger/releases/latest)
* [Source code on GitHub](https://github.com/PNNL-Comp-Mass-Spec/MASIC-Results-Merger)

### Acknowledgment

All publications that utilize this software should provide appropriate acknowledgement to PNNL and the MASIC-Results-Merger GitHub repository. However, if the software is extended or modified, then any subsequent publications should include a more extensive statement, as shown in the Readme file for the given application or on the website that more fully describes the application.

### Disclaimer

These programs are primarily designed to run on Windows machines. Please use them at your own risk. This material was prepared as an account of work sponsored by an agency of the United States Government. Neither the United States Government nor the United States Department of Energy, nor Battelle, nor any of their employees, makes any warranty, express or implied, or assumes any legal liability or responsibility for the accuracy, completeness, or usefulness or any information, apparatus, product, or process disclosed, or represents that its use would not infringe privately owned rights.

Portions of this research were supported by the NIH National Center for Research Resources (Grant RR018522), the W.R. Wiley Environmental Molecular Science Laboratory (a national scientific user facility sponsored by the U.S. Department of Energy's Office of Biological and Environmental Research and located at PNNL), and the National Institute of Allergy and Infectious Diseases (NIH/DHHS through interagency agreement Y1-AI-4894-01). PNNL is operated by Battelle Memorial Institute for the U.S. Department of Energy under contract DE-AC05-76RL0 1830.

We would like your feedback about the usefulness of the tools and information provided by the Resource. Your suggestions on how to increase their value to you will be appreciated. Please e-mail any comments to proteomics@pnl.gov
