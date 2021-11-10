using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PRISM;
using PRISM.FileProcessor;

namespace MASICResultsMerger
{
    /// <summary>
    /// This program merges the contents of a tab-delimited peptide hit results file from PHRP
    /// (for MS-GF+, MaxQuant, X!Tandem, etc.) with the corresponding MASIC results files,
    /// appending the relevant MASIC stats to each peptide hit result
    /// </summary>
    public static class Program
    {
        // Ignore Spelling: fht, Mage, tsv

        private const string PROGRAM_DATE = "November 9, 2021";

        private static string mInputFilePath;
        private static bool mCreateDartIdInputFile;
        private static bool mMageResults;
        private static bool mMergeWildcardResults;

        private static string mMASICResultsDirectoryPath;
        private static string mOutputDirectoryPath;

        private static string mOutputDirectoryAlternatePath;
        private static bool mRecreateDirectoryHierarchyInAlternatePath;

        private static bool mRecurseDirectories;
        private static int mRecurseDirectoriesMaxLevels;

        private static bool mLogMessagesToFile;

        private static int mScanNumberColumn;
        private static bool mSeparateByCollisionMode;

        private static bool mTraceMode;

        private static MASICResultsMerger mMASICResultsMerger;
        private static DateTime mLastProgressReportTime;
        private static int mLastProgressReportValue;

        private static int Main()
        {
            // Returns 0 if no error, error code if an error
            var commandLineParser = new clsParseCommandLine();
            mInputFilePath = string.Empty;
            mCreateDartIdInputFile = false;
            mMageResults = false;
            mMergeWildcardResults = false;
            mMASICResultsDirectoryPath = string.Empty;
            mOutputDirectoryPath = string.Empty;
            mRecurseDirectories = false;
            mRecurseDirectoriesMaxLevels = 0;
            mLogMessagesToFile = false;
            mScanNumberColumn = MASICResultsMerger.DEFAULT_SCAN_NUMBER_COLUMN;
            mSeparateByCollisionMode = false;
            mTraceMode = false;

            try
            {
                var proceed = false;
                if (commandLineParser.ParseCommandLine())
                {
                    if (SetOptionsUsingCommandLineParameters(commandLineParser))
                    {
                        proceed = true;
                    }
                }

                if (!proceed ||
                    commandLineParser.NeedToShowHelp ||
                    commandLineParser.ParameterCount + commandLineParser.NonSwitchParameterCount == 0 ||
                    mInputFilePath.Length == 0)
                {
                    ShowProgramHelp();
                    return -1;
                }

                // Note: If a parameter file is defined, settings in that file will override the options defined here
                mMASICResultsMerger = new MASICResultsMerger
                {
                    LogMessagesToFile = mLogMessagesToFile,
                    MASICResultsDirectoryPath = mMASICResultsDirectoryPath,
                    ScanNumberColumn = mScanNumberColumn,
                    SeparateByCollisionMode = mSeparateByCollisionMode,
                    CreateDartIdInputFile = mCreateDartIdInputFile,
                    MageResults = mMageResults,
                    TraceMode = mTraceMode
                };

                mMASICResultsMerger.ErrorEvent += MASICResultsMerger_ErrorEvent;
                mMASICResultsMerger.WarningEvent += MASICResultsMerger_WarningEvent;
                mMASICResultsMerger.StatusEvent += MASICResultsMerger_StatusEvent;
                mMASICResultsMerger.DebugEvent += MASICResultsMerger_DebugEvent;
                mMASICResultsMerger.ProgressUpdate += MASICResultsMerger_ProgressUpdate;
                mMASICResultsMerger.ProgressReset += MASICResultsMerger_ProgressReset;

                int returnCode;
                if (mRecurseDirectories)
                {
                    if (mMASICResultsMerger.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputDirectoryPath,
                                                                              mOutputDirectoryAlternatePath,
                                                                              mRecreateDirectoryHierarchyInAlternatePath, string.Empty,
                                                                              mRecurseDirectoriesMaxLevels))
                    {
                        returnCode = 0;
                    }
                    else
                    {
                        returnCode = (int)mMASICResultsMerger.ErrorCode;
                    }
                }
                else if (mMASICResultsMerger.ProcessFilesWildcard(mInputFilePath, mOutputDirectoryPath))
                {
                    returnCode = 0;
                }
                else
                {
                    returnCode = (int)mMASICResultsMerger.ErrorCode;
                    if (returnCode != 0)
                    {
                        ShowErrorMessage("Error while processing: " + mMASICResultsMerger.GetErrorMessage());
                    }
                }

                if (mMergeWildcardResults && mMASICResultsMerger.ProcessedDatasets.Count > 0)
                {
                    mMASICResultsMerger.MergeProcessedDatasets();
                }

                if (mLastProgressReportValue > 0)
                {
                    DisplayProgressPercent(mLastProgressReportValue, true);
                }

                return returnCode;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error occurred in modMain->Main: ", ex);
                return -1;
            }
        }

        private static void DisplayProgressPercent(int percentComplete, bool addCarriageReturn)
        {
            if (addCarriageReturn)
            {
                Console.WriteLine();
            }

            if (percentComplete > 100)
            {
                percentComplete = 100;
            }

            Console.Write("Processing: {0:N2}%", percentComplete);
            if (addCarriageReturn)
            {
                Console.WriteLine();
            }
        }

        private static string GetAppVersion()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version + " (" + PROGRAM_DATE + ")";
        }

        private static bool SetOptionsUsingCommandLineParameters(clsParseCommandLine commandLineParser)
        {
            // Returns True if no problems; otherwise, returns false

            var validParameters = new List<string>
            {
                "I",
                "M",
                "O",
                "N",
                "C",
                "Mage",
                "Append",
                "DartID",
                "S",
                "A",
                "R",
                "Trace"
            };

            try
            {
                // Make sure no invalid parameters are present
                if (commandLineParser.InvalidParametersPresent(validParameters))
                {
                    ShowErrorMessage("Invalid command line parameters",
                                     (from item in commandLineParser.InvalidParameters(validParameters) select "/" + item).ToList());
                    return false;
                }

                // Query commandLineParser to see if various parameters are present
                if (commandLineParser.RetrieveValueForParameter("I", out var inputFile))
                {
                    mInputFilePath = inputFile;
                }
                else if (commandLineParser.NonSwitchParameterCount > 0)
                {
                    mInputFilePath = commandLineParser.RetrieveNonSwitchParameter(0);
                }

                if (commandLineParser.RetrieveValueForParameter("M", out var masicResultsDir))
                {
                    mMASICResultsDirectoryPath = masicResultsDir;
                }

                if (commandLineParser.RetrieveValueForParameter("O", out var outputDirectory))
                {
                    mOutputDirectoryPath = outputDirectory;
                }

                if (commandLineParser.RetrieveValueForParameter("N", out var scanNumColumnIndex))
                {
                    if (int.TryParse(scanNumColumnIndex, out var value))
                    {
                        mScanNumberColumn = value;
                    }
                }

                if (commandLineParser.IsParameterPresent("C"))
                {
                    mSeparateByCollisionMode = true;
                }

                if (commandLineParser.IsParameterPresent("Mage"))
                {
                    mMageResults = true;
                }

                if (commandLineParser.IsParameterPresent("Append"))
                {
                    mMergeWildcardResults = true;
                }

                if (commandLineParser.IsParameterPresent("DartID"))
                {
                    mCreateDartIdInputFile = true;
                }

                if (commandLineParser.RetrieveValueForParameter("S", out var maxLevelsToRecurse))
                {
                    mRecurseDirectories = true;
                    if (int.TryParse(maxLevelsToRecurse, out var levels))
                    {
                        mRecurseDirectoriesMaxLevels = levels;
                    }
                }

                if (commandLineParser.RetrieveValueForParameter("A", out var alternatePath))
                {
                    mOutputDirectoryAlternatePath = alternatePath;
                }

                if (commandLineParser.IsParameterPresent("R"))
                {
                    mRecreateDirectoryHierarchyInAlternatePath = true;
                }

                if (commandLineParser.IsParameterPresent("Trace"))
                {
                    mTraceMode = true;
                }
                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error parsing the command line parameters", ex);
            }

            return false;
        }

        private static void ShowErrorMessage(string message, Exception ex = null)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }

        private static void ShowErrorMessage(string message, IEnumerable<string> items)
        {
            ConsoleMsgUtils.ShowErrors(message, items);
        }

        private static void ShowProgramHelp()
        {
            try
            {
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "This program merges the contents of a tab-delimited peptide hit results file from PHRP" +
                                      "(for MS-GF+, MaxQuant, X!Tandem, etc.) with the corresponding MASIC results files, " +
                                      "appending the relevant MASIC stats to each peptide hit result, " +
                                      "writing the merged data to a new tab-delimited text file."));
                Console.WriteLine();
                Console.WriteLine("It also supports TSV files, e.g. as created by the MzidToTsvConverter");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "If the input directory includes a MASIC _ReporterIons.txt file, " +
                                      "the reporter ion intensities will also be included in the new text file."));
                Console.WriteLine();
                Console.WriteLine("Program syntax:"
                                + Environment.NewLine + Path.GetFileName(ProcessFilesOrDirectoriesBase.GetAppPath()));
                Console.WriteLine(" InputFilePathSpec [/M:MASICResultsDirectoryPath] [/O:OutputDirectoryPath]");
                Console.WriteLine(" [/N:ScanNumberColumn] [/C] [/Mage] [/Append]");
                Console.WriteLine(" [/DartID]");
                Console.WriteLine(" [/S:[MaxLevel]] [/A:AlternateOutputDirectoryPath] [/R]");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "The input file should be a tab-delimited file where one column has scan numbers. " +
                                      "By default, this program assumes the second column has scan number, but the " +
                                      "/N switch can be used to change this (see below)."));
                Console.WriteLine();
                Console.WriteLine("Common input files are:");
                Console.WriteLine("- Peptide Hit Results Processor (https://github.com/PNNL-Comp-Mass-Spec/PHRP) tab-delimited files");
                Console.WriteLine("  - MS-GF+ syn/fht file (_msgfplus_syn.txt or _msgfplus_fht.txt)");
                Console.WriteLine("  - SEQUEST Synopsis or First-Hits file (_syn.txt or _fht.txt)");
                Console.WriteLine("  - XTandem _xt.txt file");
                Console.WriteLine("- MzidToTSVConverter (https://github.com/PNNL-Comp-Mass-Spec/Mzid-To-Tsv-Converter) .TSV files");
                Console.WriteLine("  - This is a tab-delimited text file created from a .mzid file (e.g. from MS-GF+)");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "If the MASIC result files are not in the same directory as the input file, " +
                                      "use /M to define the path to the correct directory."));
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "The output directory switch is optional. " +
                                      "If omitted, the output file will be created in the same directory as the input file"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Use /N to change the column number that contains scan number in the input file. " +
                                      "The default is 2 (meaning /N:2)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "When reading data with _ReporterIons.txt files, you can use /C to specify " +
                                      "that a separate output file be created for each collision mode type " + "in the input file (typically PQD, CID, and ETD)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Use /Mage to specify that the input file is a results file from Mage Extractor. " +
                                      "This file will contain results from several analysis jobs; the first column " +
                                      "in this file must be Job and the remaining columns must be the standard " +
                                      "Synopsis or First-Hits columns supported by PHRPReader. " +
                                      "In addition, the input directory must have a file named InputFile_metadata.txt " +
                                      "(this file will have been auto-created by Mage Extractor)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Use /Append to merge results from multiple datasets together as a single file; " +
                                      "this is only applicable when the InputFilePathSpec includes a * wildcard and multiple files are matched. " +
                                      "The merged results file will have DatasetID values of 1, 2, 3, etc. " +
                                      "along with a second file mapping DatasetID to Dataset Name"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Use /DartID to only list each peptide once per scan. " +
                                      "The Protein column will list the first protein, while the " +
                                      "Proteins column will be a comma separated list of all of the proteins. " +
                                      "This format is compatible with DART-ID (https://www.ncbi.nlm.nih.gov/pubmed/31260443)"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Use /S to process all valid files in the input directory and subdirectories. " +
                                      "Include a number after /S (like /S:2) to limit the level of subdirectories to examine. " +
                                      "When using /S, you can redirect the output of the results using /A to specify an alternate output directory. " +
                                      "When using /S, you can use /R to re-create the input directory hierarchy in the alternate output directory (if defined)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                                      "Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)"));
                Console.WriteLine("Version: " + GetAppVersion());
                Console.WriteLine();
                Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov");
                Console.WriteLine("Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics");
                Console.WriteLine();

                // Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
                System.Threading.Thread.Sleep(750);
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error displaying the program syntax: " + ex.Message);
            }
        }

        private static void MASICResultsMerger_ErrorEvent(string message, Exception ex)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }

        private static void MASICResultsMerger_WarningEvent(string message)
        {
            ConsoleMsgUtils.ShowWarning(message);
        }

        private static void MASICResultsMerger_StatusEvent(string message)
        {
            Console.WriteLine(message);
        }

        private static void MASICResultsMerger_DebugEvent(string message)
        {
            ConsoleMsgUtils.ShowDebug(message);
        }

        private static void MASICResultsMerger_ProgressUpdate(string taskDescription, float percentComplete)
        {
            const int PERCENT_REPORT_INTERVAL = 25;
            const int PROGRESS_DOT_INTERVAL_MSEC = 250;
            if (percentComplete >= mLastProgressReportValue)
            {
                if (mMageResults)
                {
                    if (mLastProgressReportValue > 0 && mLastProgressReportValue < 100)
                    {
                        Console.WriteLine();
                        DisplayProgressPercent(mLastProgressReportValue, false);
                        Console.WriteLine();
                    }
                }
                else
                {
                    if (mLastProgressReportValue > 0)
                    {
                        Console.WriteLine();
                    }

                    DisplayProgressPercent(mLastProgressReportValue, false);
                }

                mLastProgressReportValue += PERCENT_REPORT_INTERVAL;
                mLastProgressReportTime = DateTime.UtcNow;
            }
            else if (DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC)
            {
                mLastProgressReportTime = DateTime.UtcNow;
                if (!mMageResults)
                {
                    Console.Write(".");
                }
            }
        }

        private static void MASICResultsMerger_ProgressReset()
        {
            mLastProgressReportTime = DateTime.UtcNow;
            mLastProgressReportValue = 0;
        }
    }
}
