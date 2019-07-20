using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using PHRPReader;

namespace MASICResultsMerger
{
    /// This class merges the contents of a tab-delimited peptide hit results file
    /// (e.g. from SEQUEST, X!Tandem, or MS-GF+) with the corresponding MASIC results files,
    /// appending the relevant MASIC stats for each peptide hit result
    ///
    /// -------------------------------------------------------------------------------
    /// Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
    /// Program started November 26, 2008
    ///
    /// E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
    /// Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/
    /// -------------------------------------------------------------------------------
    ///
    /// Licensed under the Apache License, Version 2.0; you may not use this file except
    /// in compliance with the License.  You may obtain a copy of the License at
    /// http://www.apache.org/licenses/LICENSE-2.0
    ///
    class MASICResultsMerger : PRISM.FileProcessor.ProcessFilesBase
    {
        public MASICResultsMerger()
        {
            mFileDate = "July 19, 2019";
            ProcessedDatasets = new List<ProcessedFileInfo>();

            InitializeLocalVariables();
        }

        #region "Constants and Enums"

        const string SIC_STATS_FILE_EXTENSION = "_SICStats.txt";
        const string SCAN_STATS_FILE_EXTENSION = "_ScanStats.txt";
        const string REPORTER_IONS_FILE_EXTENSION = "_ReporterIons.txt";

        public const string RESULTS_SUFFIX = "_PlusSICStats.txt";
        public const int DEFAULT_SCAN_NUMBER_COLUMN = 2;

        public const string SCAN_STATS_ELUTION_TIME_COLUMN = "ElutionTime";

        /// <summary>
        /// Error codes specialized for this class
        /// </summary>
        public enum eResultsProcessorErrorCodes
        {
            NoError = 0,
            MissingMASICFiles = 1,
            MissingMageFiles = 2,
            UnspecifiedError = -1,
        }

        //  ReSharper disable UnusedMember.Local
        //  ReSharper disable UnusedMember.Global
        private enum eScanStatsColumns
        {
            Dataset = 0,
            ScanNumber = 1,
            ScanTime = 2,
            ScanType = 3,
            TotalIonIntensity = 4,
            BasePeakIntensity = 5,
            BasePeakMZ = 6,
            BasePeakSignalToNoiseRatio = 7,
            IonCount = 8,
            IonCountRaw = 9,
        }

        private enum eSICStatsColumns
        {
            Dataset = 0,
            ParentIonIndex = 1,
            MZ = 2,
            SurveyScanNumber = 3,
            FragScanNumber = 4,
            OptimalPeakApexScanNumber = 5,
            PeakApexOverrideParentIonIndex = 6,
            CustomSICPeak = 7,
            PeakScanStart = 8,
            PeakScanEnd = 9,
            PeakScanMaxIntensity = 10,
            PeakMaxIntensity = 11,
            PeakSignalToNoiseRatio = 12,
            FWHMInScans = 13,
            PeakArea = 14,
            ParentIonIntensity = 15,
            PeakBaselineNoiseLevel = 16,
            PeakBaselineNoiseStDev = 17,
            PeakBaselinePointsUsed = 18,
            StatMomentsArea = 19,
            CenterOfMassScan = 20,
            PeakStDev = 21,
            PeakSkew = 22,
            PeakKSStat = 23,
            StatMomentsDataCountUsed = 24,
        }

        private enum eReporterIonStatsColumns
        {
            Dataset = 0,
            ScanNumber = 1,
            CollisionMode = 2,
            ParentIonMZ = 3,
            BasePeakIntensity = 4,
            BasePeakMZ = 5,
            ReporterIonIntensityMax = 6,
        }

        //  ReSharper restore UnusedMember.Local
        //  ReSharper restore UnusedMember.Global

        #endregion

        #region "Classwide Variables"

        private eResultsProcessorErrorCodes mLocalErrorCode;

        private string mMASICResultsDirectoryPath = string.Empty;

        #endregion

        #region "Properties"

        /// <summary>
        /// When true, create a DART-ID compatible file from the Dataset_PlusSICStats.txt file
        /// </summary>
        public bool CreateDartIdInputFile { get; set; }

        public bool MageResults { get; set; }

        public string MASICResultsDirectoryPath
        {
            get => mMASICResultsDirectoryPath ?? string.Empty;
            set => mMASICResultsDirectoryPath = value ?? string.Empty;
        }

        /// <summary>
        /// Information about the datasets processed
        /// </summary>
        public List<ProcessedFileInfo> ProcessedDatasets { get; }

        /// <summary>
        /// When true, a separate output file will be created for each collision mode type; this is only possible if a _ReporterIons.txt file exists
        /// </summary>
        public bool SeparateByCollisionMode { get; set; }

        /// <summary>
        ///  For the input file, defines which column tracks scan number; the first column is column 1 (not zero)
        /// </summary>
        public int ScanNumberColumn { get; set; }

        #endregion

        private bool FindMASICFiles(string masicResultsDirectory, DatasetInfo datasetInfo, MASICFileInfo masicFiles, string masicFileSearchInfo,
                                    int job)
        {
            var triedDatasetID = false;
            var success = false;
            try
            {
                Console.WriteLine();
                ShowMessage("Looking for MASIC data files that correspond to " + datasetInfo.DatasetName);
                var datasetName = string.Copy(datasetInfo.DatasetName);
                //  Use a loop to try various possible dataset names
                while (true)
                {
                    var scanStatsFile = new FileInfo(Path.Combine(masicResultsDirectory, datasetName + SCAN_STATS_FILE_EXTENSION));
                    var sicStatsFile = new FileInfo(Path.Combine(masicResultsDirectory, datasetName + SIC_STATS_FILE_EXTENSION));
                    var reportIonsFile = new FileInfo(Path.Combine(masicResultsDirectory, datasetName + REPORTER_IONS_FILE_EXTENSION));
                    if (scanStatsFile.Exists || sicStatsFile.Exists)
                    {
                        if (scanStatsFile.Exists)
                        {
                            masicFiles.ScanStatsFileName = scanStatsFile.Name;
                        }

                        if (sicStatsFile.Exists)
                        {
                            masicFiles.SICStatsFileName = sicStatsFile.Name;
                        }

                        if (reportIonsFile.Exists)
                        {
                            masicFiles.ReporterIonsFileName = reportIonsFile.Name;
                        }

                        success = true;
                        break;
                    }

                    //  Find the last underscore in datasetName, then remove it and any text after it
                    var charIndex = datasetName.LastIndexOf('_');
                    if (charIndex > 0)
                    {
                        datasetName = datasetName.Substring(0, charIndex);
                    }
                    else if (!triedDatasetID && datasetInfo.DatasetID > 0)
                    {
                        datasetName = datasetInfo.DatasetID + "_" + datasetInfo.DatasetName;
                        triedDatasetID = true;
                    }
                    else
                    {
                        //  No more underscores; we're unable to determine the dataset name
                        break;
                    }

                }

            }
            catch (Exception ex)
            {
                HandleException("Error in FindMASICFiles", ex);
            }

            if (!success)
            {
                ShowMessage("  Error: Unable to find the MASIC data files " + masicFileSearchInfo);
                if (job != 0)
                {
                    ShowMessage("         Job " + job + " will not have MASIC results");
                }

                return false;
            }

            if (string.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName) && string.IsNullOrWhiteSpace(masicFiles.SICStatsFileName))
            {
                ShowMessage("  Error: the SIC stats and/or scan stats files were not found " + masicFileSearchInfo);
                if (job != 0)
                {
                    ShowMessage("         Job " + job + " will not have MASIC results");
                }

                return false;
            }

            if (string.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName))
            {
                ShowMessage("  Note: The MASIC SIC stats file was found, but the ScanStats file dose not exist " + masicFileSearchInfo);
            }
            else if (string.IsNullOrWhiteSpace(masicFiles.SICStatsFileName))
            {
                ShowMessage("  Note: The MASIC ScanStats file was found, but the SIC stats file dose not exist " + masicFileSearchInfo);
            }

            return true;
        }

        private void FindScanNumColumn(FileSystemInfo inputFile, IList<string> lineParts)
        {
            //  Check for a column named "ScanNum" or "ScanNumber" or "Scan Num" or "Scan Number"
            //  If found, override ScanNumberColumn
            for (var colIndex = 0; colIndex < lineParts.Count; colIndex++)
            {
                if (lineParts[colIndex].Equals("Scan", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("ScanNum", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("Scan Num", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("ScanNumber", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("Scan Number", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("Scan#", StringComparison.OrdinalIgnoreCase) ||
                    lineParts[colIndex].Equals("Scan #", StringComparison.OrdinalIgnoreCase))
                {
                    if (ScanNumberColumn != colIndex + 1)
                    {
                        ScanNumberColumn = colIndex + 1;
                        ShowMessage(string.Format("Note: Reading scan numbers from column {0} ({1}) in file {2}",
                                                  ScanNumberColumn, lineParts[colIndex], inputFile.Name));
                        break;
                    }

                }

            }

        }

        private List<string> GetScanStatsHeaders()
        {
            var scanStatsColumns = new List<string>()
            {
                SCAN_STATS_ELUTION_TIME_COLUMN,
                "ScanType",
                "TotalIonIntensity",
                "BasePeakIntensity",
                "BasePeakMZ"
            };
            return scanStatsColumns;
        }

        private List<string> GetSICStatsHeaders()
        {
            var sicStatsColumns = new List<string>
            {
                "Optimal_Scan_Number",
                "PeakMaxIntensity",
                "PeakSignalToNoiseRatio",
                "FWHMInScans",
                "PeakArea",
                "ParentIonIntensity",
                "ParentIonMZ",
                "StatMomentsArea",
            };
            return sicStatsColumns;
        }

        private string FlattenList(IReadOnlyList<string> lstData)
        {
            return string.Join("\t", lstData);
        }

        private string FlattenArray(IList<string> lineParts, int indexStart)
        {
            var text = string.Empty;
            if (lineParts == null || lineParts.Count <= 0)
            {
                return text;
            }

            for (var index = indexStart; index < lineParts.Count; index++)
            {
                string column;
                if (lineParts[index] == null)
                {
                    column = string.Empty;
                }
                else
                {
                    column = string.Copy(lineParts[index]);
                }

                if (index > indexStart)
                {
                    text += '\t' + column;
                }
                else
                {
                    text = string.Copy(column);
                }

            }

            return text;
        }


        /// <summary>
        /// Get the current error state, if any
        /// </summary>
        /// <returns>Returns an empty string if no error</returns>
        public override string GetErrorMessage()
        {
            string errorMessage;
            if (ErrorCode == ProcessFilesErrorCodes.LocalizedError ||
                ErrorCode == ProcessFilesErrorCodes.NoError)
            {
                switch (mLocalErrorCode)
                {
                    case eResultsProcessorErrorCodes.NoError:
                        errorMessage = string.Empty;
                        break;
                    case eResultsProcessorErrorCodes.UnspecifiedError:
                        errorMessage = "Unspecified localized error";
                        break;
                    case eResultsProcessorErrorCodes.MissingMASICFiles:
                        errorMessage = "Missing MASIC Files";
                        break;
                    case eResultsProcessorErrorCodes.MissingMageFiles:
                        errorMessage = "Missing Mage Extractor Files";
                        break;
                    default:
                        errorMessage = "Unknown error state";
                        break;
                }
            }
            else
            {
                errorMessage = GetBaseClassErrorMessage();
            }

            return errorMessage;
        }

        private void InitializeLocalVariables()
        {
            mMASICResultsDirectoryPath = string.Empty;
            ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN;
            SeparateByCollisionMode = false;
            mLocalErrorCode = eResultsProcessorErrorCodes.NoError;
        }

        private bool MergePeptideHitAndMASICFiles(FileSystemInfo inputFile, string outputDirectoryPath, Dictionary<int, ScanStatsData> dctScanStats,
                                                  IReadOnlyDictionary<int, SICStatsData> dctSICStats, string reporterIonHeaders)
        {
            StreamWriter[] writers;
            int[] linesWritten;

            int outputFileCount;

            //  The Key is the collision mode and the value is the path
            KeyValuePair<string, string>[] outputFilePaths;
            string baseFileName;

            var blankAdditionalScanStatsColumns = string.Empty;
            var blankAdditionalSICColumns = string.Empty;
            var blankAdditionalReporterIonColumns = string.Empty;

            Dictionary<string, int> dctCollisionModeFileMap;

            try
            {
                if (!inputFile.Exists)
                {
                    ShowErrorMessage("File not found: " + inputFile.FullName);
                    return false;
                }

                baseFileName = Path.GetFileNameWithoutExtension(inputFile.Name);
                if (string.IsNullOrWhiteSpace(reporterIonHeaders))
                {
                    reporterIonHeaders = string.Empty;
                }

                //  Define the output file path
                if (outputDirectoryPath == null)
                {
                    outputDirectoryPath = string.Empty;
                }

                if (SeparateByCollisionMode)
                {
                    outputFilePaths = SummarizeCollisionModes(inputFile, baseFileName, outputDirectoryPath, dctScanStats, out dctCollisionModeFileMap);
                    outputFileCount = outputFilePaths.Length;
                    if (outputFileCount < 1)
                    {
                        return false;
                    }

                }
                else
                {
                    dctCollisionModeFileMap = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
                    outputFileCount = 1;
                    outputFilePaths = new KeyValuePair<string, string>[0];
                    outputFilePaths[0] = new KeyValuePair<string, string>("", Path.Combine(outputDirectoryPath, baseFileName + RESULTS_SUFFIX));
                }

                writers = new StreamWriter[outputFileCount - 1];
                linesWritten = new int[outputFileCount - 1];

                for (var index = 0; index < outputFileCount; index++)
                {
                    writers[index] =
                        new StreamWriter(new FileStream(outputFilePaths[index].Value, FileMode.Create, FileAccess.Write, FileShare.Read));
                }

            }
            catch (Exception ex)
            {
                HandleException("Error creating the merged output file", ex);
                return false;
            }

            try
            {
                Console.WriteLine();
                ShowMessage("Parsing " + inputFile.Name + " and writing " + Path.GetFileName(outputFilePaths[0].Value));
                if (ScanNumberColumn < 1)
                {
                    //  Assume the scan number is in the second column
                    ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN;
                }

                //  Read from reader and write out to the file(s) in swOutFile
                using (var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    var linesRead = 0;
                    var writeReporterIonStats = false;
                    var writeSICStats = (dctSICStats.Count > 0);
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        var collisionModeCurrentScan = string.Empty;
                        if (string.IsNullOrWhiteSpace(dataLine))
                        {
                            continue;
                        }

                        linesRead++;
                        var lineParts = dataLine.Split('\t').ToList();
                        if (linesRead == 1)
                        {
                            string headerLine;
                            //  Write out an updated header line
                            if (lineParts.Count >= ScanNumberColumn && int.TryParse(lineParts[ScanNumberColumn - 1], out _))
                            {
                                //  The input file doesn't have a header line; we will add one, using generic column names for the data in the input file
                                var genericHeaders = new List<string>();
                                for (var index = 0; index < lineParts.Count; index++)
                                {
                                    genericHeaders.Add("Column" + index.ToString("00"));
                                }

                                headerLine = FlattenList(genericHeaders);
                            }
                            else
                            {
                                //  The input file does have a text-based header
                                headerLine = string.Copy(dataLine);
                                FindScanNumColumn(inputFile, lineParts);
                                //  Clear splitLine so that this line gets skipped
                                lineParts.Clear();
                            }

                            var scanStatsHeaders = GetScanStatsHeaders();
                            var sicStatsHeaders = GetSICStatsHeaders();
                            if (!writeSICStats)
                            {
                                sicStatsHeaders.Clear();
                            }

                            //  Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
                            blankAdditionalScanStatsColumns = new string('\t', scanStatsHeaders.Count - 1);
                            if (writeSICStats)
                            {
                                blankAdditionalSICColumns = new string('\t', sicStatsHeaders.Count);
                            }

                            //  Initialize blankAdditionalReporterIonColumns
                            if (reporterIonHeaders.Length > 0)
                            {
                                blankAdditionalReporterIonColumns =
                                    new string('\t', reporterIonHeaders.Split('\t').ToList().Count - 1);
                            }

                            //  Initialize the AddOn header columns
                            var addonHeaders = FlattenList(scanStatsHeaders);
                            if (writeSICStats)
                            {
                                addonHeaders += '\t' + FlattenList(sicStatsHeaders);
                            }

                            if (reporterIonHeaders.Length > 0)
                            {
                                //  Append the reporter ion stats columns
                                addonHeaders += '\t' + reporterIonHeaders;
                                writeReporterIonStats = true;
                            }

                            //  Write out the headers
                            for (var index = 0; index < outputFileCount; index++)
                            {
                                writers[index].WriteLine(headerLine + '\t' + addonHeaders);
                            }

                        }

                        if (lineParts.Count < ScanNumberColumn ||
                            !int.TryParse(lineParts[ScanNumberColumn - 1], out var scanNumber))
                        {
                            continue;
                        }

                        //  Look for scanNumber in dctScanStats
                        var addonColumns = new List<string>();
                        if (!dctScanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                        {
                            //  Match not found; use the blank columns in blankAdditionalScanStatsColumns
                            addonColumns.Add(blankAdditionalScanStatsColumns);
                        }
                        else
                        {
                            addonColumns.Add(scanStatsEntry.ElutionTime);
                            addonColumns.Add(scanStatsEntry.ScanType);
                            addonColumns.Add(scanStatsEntry.TotalIonIntensity);
                            addonColumns.Add(scanStatsEntry.BasePeakIntensity);
                            addonColumns.Add(scanStatsEntry.BasePeakMZ);
                        }

                        if (writeSICStats)
                        {
                            if (!dctSICStats.TryGetValue(scanNumber, out var sicStatsEntry))
                            {
                                //  Match not found; use the blank columns in blankAdditionalSICColumns
                                addonColumns.Add(blankAdditionalSICColumns);
                            }
                            else
                            {
                                addonColumns.Add(sicStatsEntry.OptimalScanNumber);
                                addonColumns.Add(sicStatsEntry.PeakMaxIntensity);
                                addonColumns.Add(sicStatsEntry.PeakSignalToNoiseRatio);
                                addonColumns.Add(sicStatsEntry.FWHMInScans);
                                addonColumns.Add(sicStatsEntry.PeakArea);
                                addonColumns.Add(sicStatsEntry.ParentIonIntensity);
                                addonColumns.Add(sicStatsEntry.ParentIonMZ);
                                addonColumns.Add(sicStatsEntry.StatMomentsArea);
                            }

                        }

                        if (writeReporterIonStats)
                        {
                            if (scanStatsEntry == null || string.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode))
                            {
                                //  Collision mode is not defined; append blank columns
                                addonColumns.Add(string.Empty);
                                addonColumns.Add(blankAdditionalReporterIonColumns);
                            }
                            else
                            {
                                //  Collision mode is defined
                                addonColumns.Add(scanStatsEntry.CollisionMode);
                                addonColumns.Add(scanStatsEntry.ReporterIonData);
                                collisionModeCurrentScan = string.Copy(scanStatsEntry.CollisionMode);
                            }

                        }
                        else if (SeparateByCollisionMode)
                        {
                            if (scanStatsEntry == null)
                            {
                                collisionModeCurrentScan = string.Empty;
                            }
                            else
                            {
                                collisionModeCurrentScan = string.Copy(scanStatsEntry.CollisionMode);
                            }

                        }

                        var outFileIndex = 0;
                        if (SeparateByCollisionMode && outputFileCount > 1)
                        {
                            if (collisionModeCurrentScan != null)
                            {
                                //  Determine the correct output file
                                if (!dctCollisionModeFileMap.TryGetValue(collisionModeCurrentScan, out outFileIndex))
                                {
                                    outFileIndex = 0;
                                }

                            }

                        }

                        writers[outFileIndex].WriteLine(dataLine + '\t' + FlattenList(addonColumns));
                        linesWritten[outFileIndex]++;
                    }
                }

                //  Close the output files
                if (writers != null)
                {
                    for (var index = 0; index < outputFileCount; index++)
                    {
                        if (writers[index] != null)
                        {
                            writers[index].Close();
                        }

                    }

                }

                if (CreateDartIdInputFile)
                {
                    var preprocessor = new DartIdPreprocessor();

                    foreach (var item in outputFilePaths.ToList())
                    {
                        var successConsolidating = preprocessor.ConsolidatePSMs(item.Value, false);
                    }

                }

                //  See if any of the files had no data written to them
                //  If there are, then delete the empty output file
                //  However, retain at least one output file
                var emptyOutFileCount = 0;
                for (var index = 0; index < outputFileCount; index++)
                {
                    if (linesWritten[index] == 0)
                    {
                        emptyOutFileCount++;
                    }

                }

                var outputPathEntry = new ProcessedFileInfo(baseFileName);
                if (emptyOutFileCount == 0)
                {
                    foreach (var item in outputFilePaths.ToList())
                    {
                        outputPathEntry.AddOutputFile(item.Key, item.Value);
                    }

                }
                else
                {
                    if (emptyOutFileCount == outputFileCount)
                    {
                        //  All output files are empty
                        //  Pretend the first output file actually contains data
                        linesWritten[0] = 1;
                    }

                    for (var index = 0; index < outputFileCount; index++)
                    {
                        //  Wait 250 msec before continuing
                        Thread.Sleep(250);
                        if (linesWritten[index] == 0)
                        {
                            try
                            {
                                ShowMessage("Deleting empty output file: " + Environment.NewLine +
                                            " --> " + Path.GetFileName(outputFilePaths[index].Value));
                                File.Delete(outputFilePaths[index].Value);
                            }
                            catch (Exception)
                            {
                                //  Ignore errors here
                            }

                        }
                        else
                        {
                            outputPathEntry.AddOutputFile(outputFilePaths[index].Key, outputFilePaths[index].Value);
                        }

                    }

                }

                if (outputPathEntry.OutputFiles.Count > 0)
                {
                    ProcessedDatasets.Add(outputPathEntry);
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in MergePeptideHitAndMASICFiles", ex);
                return false;
            }

        }

        public bool MergeProcessedDatasets()
        {
            try
            {
                if (ProcessedDatasets.Count == 1)
                {
                    //  Nothing to merge
                    ShowMessage("Only one dataset has been processed by the MASICResultsMerger; nothing to merge");
                    return true;
                }

                //  Determine the base filename and collision modes used
                var baseFileName = string.Empty;
                var collisionModes = new SortedSet<string>();
                var datasetNameIdMap = new Dictionary<string, int>();
                foreach (var processedDataset in ProcessedDatasets)
                {
                    foreach (var processedFile in processedDataset.OutputFiles)
                    {
                        if (!collisionModes.Contains(processedFile.Key))
                        {
                            collisionModes.Add(processedFile.Key);
                        }

                    }

                    if (!datasetNameIdMap.ContainsKey(processedDataset.BaseName))
                    {
                        datasetNameIdMap.Add(processedDataset.BaseName, datasetNameIdMap.Count + 1);
                    }

                    //  Find the characters common to all of the processed datasets
                    var candidateName = processedDataset.BaseName;
                    if (string.IsNullOrEmpty(baseFileName))
                    {
                        baseFileName = string.Copy(candidateName);
                    }
                    else
                    {
                        var charsInCommon = 0;
                        for (var index = 0; index < baseFileName.Length; index++)
                        {
                            if (index >= candidateName.Length)
                            {
                                break;
                            }

                            if (candidateName.ToLower()[index] != baseFileName.ToLower()[index])
                            {
                                break;
                            }

                            charsInCommon++;
                        }

                        if (charsInCommon > 1)
                        {
                            baseFileName = baseFileName.Substring(0, charsInCommon);
                            //  Possibly backtrack to the previous underscore
                            var lastUnderscore = baseFileName.LastIndexOf("_", StringComparison.Ordinal);
                            if (lastUnderscore >= 4)
                            {
                                baseFileName = baseFileName.Substring(0, lastUnderscore);
                            }

                        }

                    }

                }

                if (collisionModes.Count == 0)
                {
                    ShowErrorMessage("None of the processed datasets had any output files");
                }

                baseFileName = "MergedData_" + baseFileName;
                //  Open the output files
                var outputFileHandles = new Dictionary<string, StreamWriter>();
                var outputFileHeaderWritten = new Dictionary<string, bool>();
                foreach (var collisionMode in collisionModes)
                {
                    string outputFileName;
                    if (collisionMode == ProcessedFileInfo.COLLISION_MODE_NOT_DEFINED)
                    {
                        outputFileName = baseFileName + RESULTS_SUFFIX;
                    }
                    else
                    {
                        outputFileName = baseFileName + "_" + collisionMode + RESULTS_SUFFIX;
                    }

                    outputFileHandles.Add(collisionMode,
                                          new StreamWriter(new FileStream(Path.Combine(mOutputDirectoryPath, outputFileName), FileMode.Create,
                                                                          FileAccess.Write)));

                    outputFileHeaderWritten.Add(collisionMode, false);
                }

                //  Create the DatasetMap file
                using (var writer = new StreamWriter(new FileStream(Path.Combine(mOutputDirectoryPath, baseFileName + "_DatasetMap.txt"),
                                                                    FileMode.Create, FileAccess.Write)))
                {
                    writer.WriteLine("DatasetID" + '\t' + "DatasetName");
                    foreach (var datasetMapping in datasetNameIdMap)
                    {
                        writer.WriteLine(datasetMapping.Value + '\t' + datasetMapping.Key);
                    }
                }

                //  Merge the files
                foreach (var processedDataset in ProcessedDatasets)
                {
                    var datasetId = datasetNameIdMap[processedDataset.BaseName];
                    foreach (var sourceFile in processedDataset.OutputFiles)
                    {
                        var collisionMode = sourceFile.Key;

                        if (!outputFileHandles.TryGetValue(collisionMode, out var writer))
                        {
                            Console.WriteLine("Warning: unrecognized collision mode; skipping " + sourceFile.Value);
                            continue;
                        }

                        if (!File.Exists(sourceFile.Value))
                        {
                            Console.WriteLine("Warning: input file not found; skipping " + sourceFile.Value);
                            continue;
                        }

                        using (var reader = new StreamReader(new FileStream(sourceFile.Value, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                        {
                            var linesRead = 0;
                            while (!reader.EndOfStream)
                            {
                                var dataLine = reader.ReadLine();
                                linesRead++;
                                if (linesRead == 1)
                                {
                                    if (outputFileHeaderWritten[collisionMode])
                                    {
                                        //  skip this line
                                        continue;
                                    }

                                    writer.WriteLine("DatasetID" + '\t' + dataLine);
                                    outputFileHeaderWritten[collisionMode] = true;
                                }
                                else
                                {
                                    writer.WriteLine(datasetId + '\t' + dataLine);
                                }

                            }
                        }

                    }

                }

                foreach (var outputFile in outputFileHandles)
                {
                    outputFile.Value.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in MergeProcessedDatasets", ex);
                return false;
            }

        }

        /// <summary>
        /// Main processing function
        /// </summary>
        /// <param name="inputFilePath">Input file path</param>
        /// <param name="outputDirectoryPath">Output directory path</param>
        /// <param name="parameterFilePath">Parameter file path (Ignored)</param>
        /// <param name="resetErrorCode">If true, reset the error code</param>
        /// <returns>True if success, False if failure</returns>
        public override bool ProcessFile(string inputFilePath, string outputDirectoryPath, string parameterFilePath, bool resetErrorCode)
        {
            bool success;
            string masicResultsDirectory;
            if (resetErrorCode)
            {
                SetLocalErrorCode(eResultsProcessorErrorCodes.NoError);
            }

            if (string.IsNullOrEmpty(inputFilePath))
            {
                ShowMessage("Input file name is empty");
                SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath);
                return false;
            }

            //  Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
            if (!CleanupFilePaths(ref inputFilePath, ref outputDirectoryPath))
            {
                SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError);
                return false;
            }

            var fiInputFile = new FileInfo(inputFilePath);
            if (string.IsNullOrWhiteSpace(mMASICResultsDirectoryPath))
            {
                masicResultsDirectory = fiInputFile.DirectoryName;
            }
            else
            {
                masicResultsDirectory = string.Copy(mMASICResultsDirectoryPath);
            }

            ProcessedDatasets.Clear();

            if (MageResults)
            {
                success = ProcessMageExtractorFile(fiInputFile, masicResultsDirectory);
            }
            else
            {
                success = ProcessSingleJobFile(fiInputFile, masicResultsDirectory);
            }

            return success;
        }

        private bool ProcessMageExtractorFile(FileInfo fiInputFile, string masicResultsDirectory)
        {
            var dctScanStats = new Dictionary<int, ScanStatsData>();
            var dctSICStats = new Dictionary<int, SICStatsData>();
            try
            {
                //  Read the Mage Metadata file
                FileInfo fiMetadataFile;
                var metadataFileName = Path.GetFileNameWithoutExtension(fiInputFile.Name) + "_metadata.txt";

                if (fiInputFile.DirectoryName != null)
                {
                    fiMetadataFile = new FileInfo(Path.Combine(fiInputFile.DirectoryName, metadataFileName));
                }
                else
                {
                    fiMetadataFile = new FileInfo(metadataFileName);
                }

                if (!fiMetadataFile.Exists)
                {
                    ShowErrorMessage("Error: Mage Metadata File not found: " + fiMetadataFile.FullName);
                    SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMageFiles);
                    return false;
                }

                //  Keys in this dictionary are the job, values are the DatasetID and DatasetName
                var dctJobToDatasetMap = ReadMageMetadataFile(fiMetadataFile.FullName);
                if (dctJobToDatasetMap == null || dctJobToDatasetMap.Count == 0)
                {
                    ShowErrorMessage("Error: ReadMageMetadataFile returned an empty job mapping");
                    return false;
                }

                string headerLine;
                int jobColumnIndex;

                // Open the Mage Extractor data file so that we can validate and cache the header row
                using (var reader = new StreamReader(new FileStream(fiInputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    headerLine = reader.ReadLine();
                    var invalidFileMessage = string.Format(
                        "Input file is not a valid Mage Extractor results file; it must contain a \"Job\" column: " +
                        fiInputFile.FullName);

                    if (string.IsNullOrWhiteSpace(headerLine))
                    {
                        ShowErrorMessage(invalidFileMessage);
                        return false;
                    }

                    var lstColumns = headerLine.Split('\t').ToList();
                    jobColumnIndex = lstColumns.IndexOf("Job");
                    if (jobColumnIndex < 0)
                    {
                        ShowErrorMessage(invalidFileMessage);
                        return false;
                    }
                }

                var scanStatsHeaders = GetScanStatsHeaders();
                var sicStatsHeaders = GetSICStatsHeaders();
                //  Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
                var blankAdditionalScanStatsColumns = new string('\t', scanStatsHeaders.Count - 1);
                var blankAdditionalSICColumns = new string('\t', sicStatsHeaders.Count);
                var outputFileName = Path.GetFileNameWithoutExtension(fiInputFile.Name) + RESULTS_SUFFIX;
                var outputFilePath = Path.Combine(mOutputDirectoryPath, outputFileName);
                var jobsSuccessfullyMerged = 0;

                //  Initialize the output file
                using (var writer = new StreamWriter(new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read)))
                {
                    //  Open the Mage Extractor data file and read the data for each job
                    var phrpReader = new clsPHRPReader(fiInputFile.FullName, clsPHRPReader.ePeptideHitResultType.Unknown, false, false, false)
                    {
                        EchoMessagesToConsole = false,
                        SkipDuplicatePSMs = false
                    };

                    RegisterEvents(phrpReader);
                    if (!phrpReader.CanRead)
                    {
                        ShowErrorMessage("Aborting since PHRPReader is not ready: " + phrpReader.ErrorMessage);
                        return false;
                    }

                    var lastJob = -1;

                    var masicDataLoaded = false;
                    var headerLineWritten = false;
                    var writeReporterIonStats = false;
                    var reporterIonHeaders = string.Empty;
                    var blankAdditionalReporterIonColumns = string.Empty;

                    while (phrpReader.MoveNext())
                    {
                        var psm = phrpReader.CurrentPSM;
                        //  Parse out the job from the current line
                        var lstColumns = psm.DataLineText.Split('\t').ToList();

                        if (!int.TryParse(lstColumns[jobColumnIndex], out var job))
                        {
                            ShowMessage("Warning: Job column does not contain a job number; skipping this entry: " + psm.DataLineText);
                            continue;
                        }

                        if (job != lastJob)
                        {
                            //  New job; read and cache the MASIC data
                            masicDataLoaded = false;

                            if (!dctJobToDatasetMap.TryGetValue(job, out var datasetInfo))
                            {
                                ShowErrorMessage("Error: Job " + job + " was not defined in the Metadata file; unable to determine the dataset");
                            }
                            else
                            {
                                //  Look for the corresponding MASIC files in the input directory
                                var masicFiles = new MASICFileInfo();
                                var datasetNameAndDirectory = "for dataset " + datasetInfo.DatasetName + " in " + masicResultsDirectory;
                                var success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, datasetNameAndDirectory, job);
                                if (success)
                                {
                                    //  Read and cache the MASIC data
                                    dctScanStats = new Dictionary<int, ScanStatsData>();
                                    dctSICStats = new Dictionary<int, SICStatsData>();

                                    masicDataLoaded = ReadMASICData(masicResultsDirectory, masicFiles, dctScanStats, dctSICStats, out reporterIonHeaders);
                                    if (masicDataLoaded)
                                    {
                                        jobsSuccessfullyMerged++;
                                        if (jobsSuccessfullyMerged == 1)
                                        {
                                            //  Initialize blankAdditionalReporterIonColumns
                                            if (reporterIonHeaders.Length > 0)
                                            {
                                                blankAdditionalReporterIonColumns =
                                                    new string('\t', reporterIonHeaders.Split('\t').ToList().Count - 1);
                                            }

                                        }

                                    }

                                }

                            }

                        }

                        if (masicDataLoaded)
                        {
                            if (!headerLineWritten)
                            {
                                var addonHeaderColumns = FlattenList(scanStatsHeaders) + '\t' + FlattenList(sicStatsHeaders);
                                if (reporterIonHeaders.Length > 0)
                                {
                                    //  The merged data file will have reporter ion columns
                                    writeReporterIonStats = true;
                                    writer.WriteLine(headerLine + '\t' + addonHeaderColumns + '\t' + reporterIonHeaders);
                                }
                                else
                                {
                                    writer.WriteLine(headerLine + '\t' + addonHeaderColumns);
                                }

                                headerLineWritten = true;
                            }

                            //  Look for scanNumber in dctScanStats
                            var addonColumns = new List<string>();
                            if (!dctScanStats.TryGetValue(psm.ScanNumber, out var scanStatsEntry))
                            {
                                //  Match not found; use the blank columns in blankAdditionalColumns
                                addonColumns.Add(blankAdditionalScanStatsColumns);
                            }
                            else
                            {
                                addonColumns.Add(scanStatsEntry.ElutionTime);
                                addonColumns.Add(scanStatsEntry.ScanType);
                                addonColumns.Add(scanStatsEntry.TotalIonIntensity);
                                addonColumns.Add(scanStatsEntry.BasePeakIntensity);
                                addonColumns.Add(scanStatsEntry.BasePeakMZ);
                            }

                            if (!dctSICStats.TryGetValue(psm.ScanNumber, out var sicStatsEntry))
                            {
                                //  Match not found; use the blank columns in blankAdditionalSICColumns
                                addonColumns.Add(blankAdditionalSICColumns);
                            }
                            else
                            {
                                addonColumns.Add(sicStatsEntry.OptimalScanNumber);
                                addonColumns.Add(sicStatsEntry.PeakMaxIntensity);
                                addonColumns.Add(sicStatsEntry.PeakSignalToNoiseRatio);
                                addonColumns.Add(sicStatsEntry.FWHMInScans);
                                addonColumns.Add(sicStatsEntry.PeakArea);
                                addonColumns.Add(sicStatsEntry.ParentIonIntensity);
                                addonColumns.Add(sicStatsEntry.ParentIonMZ);
                                addonColumns.Add(sicStatsEntry.StatMomentsArea);
                            }

                            if (writeReporterIonStats)
                            {
                                if (scanStatsEntry == null || string.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode))
                                {
                                    //  Collision mode is not defined; append blank columns
                                    addonColumns.Add(string.Empty);
                                    addonColumns.Add(blankAdditionalReporterIonColumns);
                                }
                                else
                                {
                                    //  Collision mode is defined
                                    addonColumns.Add(scanStatsEntry.CollisionMode);
                                    addonColumns.Add(scanStatsEntry.ReporterIonData);
                                }

                            }

                            writer.WriteLine(psm.DataLineText + '\t' + FlattenList(addonColumns));
                        }
                        else
                        {
                            var blankAddonColumns = '\t' + blankAdditionalScanStatsColumns + '\t' + blankAdditionalSICColumns;

                            if (writeReporterIonStats)
                            {
                                writer.WriteLine(psm.DataLineText + blankAddonColumns + '\t' + '\t' + blankAdditionalReporterIonColumns);
                            }
                            else
                            {
                                writer.WriteLine(psm.DataLineText + blankAddonColumns);
                            }

                        }

                        UpdateProgress("Loading data from " + fiInputFile.Name, phrpReader.PercentComplete);
                        lastJob = job;
                    }
                }

                if (jobsSuccessfullyMerged > 0)
                {
                    Console.WriteLine();
                    ShowMessage("Merged MASIC results for " + jobsSuccessfullyMerged + " jobs");
                }

                if (CreateDartIdInputFile)
                {
                    var preprocessor = new DartIdPreprocessor();
                    var successConsolidating = preprocessor.ConsolidatePSMs(outputFilePath, true);
                }

                return jobsSuccessfullyMerged > 0;
            }
            catch (Exception ex)
            {
                HandleException("Error in ProcessMageExtractorFile", ex);
                return false;
            }
        }

        private bool ProcessSingleJobFile(FileSystemInfo fiInputFile, string masicResultsDirectory)
        {
            try
            {
                var datasetName = Path.GetFileNameWithoutExtension(fiInputFile.FullName);
                var datasetInfo = new DatasetInfo(datasetName, 0);
                //  Note that FindMASICFiles will first try the full filename, and if it doesn't find a match,
                //  it will start removing text from the end of the filename by looking for underscores
                //  Look for the corresponding MASIC files in the input directory
                var masicFiles = new MASICFileInfo();
                var masicFileSearchInfo = " in " + masicResultsDirectory;
                var success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, masicFileSearchInfo, 0);
                if (!success)
                {
                    SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles);
                    return false;
                }

                //  Read and cache the MASIC data
                var dctScanStats = new Dictionary<int, ScanStatsData>();
                var dctSICStats = new Dictionary<int, SICStatsData>();

                success = ReadMASICData(masicResultsDirectory, masicFiles, dctScanStats, dctSICStats, out var reporterIonHeaders);
                if (success)
                {
                    //  Merge the MASIC data with the input file
                    success = MergePeptideHitAndMASICFiles(fiInputFile, mOutputDirectoryPath, dctScanStats, dctSICStats, reporterIonHeaders);
                }

                if (success)
                {
                    ShowMessage(string.Empty, false);
                }
                else
                {
                    SetLocalErrorCode(eResultsProcessorErrorCodes.UnspecifiedError);
                    ShowErrorMessage("Error");
                }

                return success;
            }
            catch (Exception ex)
            {
                HandleException("Error in ProcessSingleJobFile", ex);
                return false;
            }

        }

        private bool ReadMASICData(
            string sourceDirectory,
            MASICFileInfo masicFiles,
            IDictionary<int, ScanStatsData> dctScanStats,
            IDictionary<int, SICStatsData> dctSICStats,
            out string reporterIonHeaders)
        {
            try
            {
                bool scanStatsRead;
                bool sicStatsRead;
                if (string.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName))
                {
                    scanStatsRead = false;
                }
                else
                {
                    scanStatsRead = ReadScanStatsFile(sourceDirectory, masicFiles.ScanStatsFileName, dctScanStats);
                }

                if (string.IsNullOrWhiteSpace(masicFiles.SICStatsFileName))
                {
                    sicStatsRead = false;
                }
                else
                {
                    sicStatsRead = ReadSICStatsFile(sourceDirectory, masicFiles.SICStatsFileName, dctSICStats);
                }

                if (string.IsNullOrWhiteSpace(masicFiles.ReporterIonsFileName))
                {
                    reporterIonHeaders = string.Empty;
                }
                else
                {
                    ReadReporterIonStatsFile(sourceDirectory, masicFiles.ReporterIonsFileName, dctScanStats, out reporterIonHeaders);
                }

                return scanStatsRead || sicStatsRead;
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadMASICData", ex);
                reporterIonHeaders = string.Empty;
                return false;
            }

        }

        private bool ReadScanStatsFile(string sourceDirectory, string scanStatsFileName, IDictionary<int, ScanStatsData> dctScanStats)
        {
            try
            {
                //  Initialize dctScanStats
                dctScanStats.Clear();
                ShowMessage("  Reading: " + scanStatsFileName);
                using (var reader = new StreamReader(new FileStream(Path.Combine(sourceDirectory, scanStatsFileName), FileMode.Open, FileAccess.Read,
                                                                    FileShare.ReadWrite)))
                {
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                        {
                            continue;
                        }

                        var lineParts = dataLine.Split('\t');
                        if (lineParts.Length < (int)eScanStatsColumns.BasePeakMZ + 1)
                        {
                            continue;
                        }

                        if (!int.TryParse(lineParts[(int)eScanStatsColumns.ScanNumber], out var scanNumber))
                        {
                            continue;
                        }

                        //  Note: the remaining values are stored as strings to prevent the number format from changing
                        var scanStatsEntry = new ScanStatsData(scanNumber)
                        {
                            ElutionTime = string.Copy(lineParts[(int)eScanStatsColumns.ScanTime]),
                            ScanType = string.Copy(lineParts[(int)eScanStatsColumns.ScanType]),
                            TotalIonIntensity = string.Copy(lineParts[(int)eScanStatsColumns.TotalIonIntensity]),
                            BasePeakIntensity = string.Copy(lineParts[(int)eScanStatsColumns.BasePeakIntensity]),
                            BasePeakMZ = string.Copy(lineParts[(int)eScanStatsColumns.BasePeakMZ]),
                            CollisionMode = string.Empty,
                            ReporterIonData = string.Empty
                        };

                        dctScanStats.Add(scanNumber, scanStatsEntry);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadScanStatsFile", ex);
                return false;
            }

        }



        private Dictionary<int, DatasetInfo> ReadMageMetadataFile(string metadataFilePath)
        {
            var dctJobToDatasetMap = new Dictionary<int, DatasetInfo>();
            var headersParsed = false;
            var jobIndex = -1;
            var datasetIndex = -1;
            var datasetIDIndex = -1;

            try
            {
                using (var reader = new StreamReader(new FileStream(metadataFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                        {
                            continue;
                        }

                        var lineParts = dataLine.Split('\t').ToList();
                        if (!headersParsed)
                        {
                            //  Look for the Job and Dataset columns
                            jobIndex = lineParts.IndexOf("Job");
                            datasetIndex = lineParts.IndexOf("Dataset");
                            datasetIDIndex = lineParts.IndexOf("Dataset_ID");

                            if (jobIndex < 0)
                            {
                                ShowErrorMessage("Job column not found in the metadata file: " + metadataFilePath);
                                return null;
                            }

                            if (datasetIndex < 0)
                            {
                                ShowErrorMessage("Dataset column not found in the metadata file: " + metadataFilePath);
                                return null;
                            }

                            if (datasetIDIndex < 0)
                            {
                                ShowErrorMessage("Dataset_ID column not found in the metadata file: " + metadataFilePath);
                                return null;
                            }

                            headersParsed = true;
                            continue;
                        }

                        if (lineParts.Count > datasetIndex)
                        {
                            if (int.TryParse(lineParts[jobIndex], out var jobNumber))
                            {
                                if (int.TryParse(lineParts[datasetIDIndex], out var datasetID))
                                {
                                    var datasetName = lineParts[datasetIndex];
                                    var datasetInfo = new DatasetInfo(datasetName, datasetID);
                                    dctJobToDatasetMap.Add(jobNumber, datasetInfo);
                                }
                                else
                                {
                                    ShowMessage("Warning: Dataset_ID number not numeric in metadata file, line " + dataLine);
                                }

                            }
                            else
                            {
                                ShowMessage("Warning: Job number not numeric in metadata file, line " + dataLine);
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadMageMetadataFile", ex);
                return null;
            }

            return dctJobToDatasetMap;
        }

        private bool ReadSICStatsFile(string sourceDirectory, string sicStatsFileName, IDictionary<int, SICStatsData> dctSICStats)
        {

            try
            {
                //  Initialize dctSICStats
                dctSICStats.Clear();
                ShowMessage("  Reading: " + sicStatsFileName);
                using (var reader = new StreamReader(new FileStream(Path.Combine(sourceDirectory, sicStatsFileName), FileMode.Open, FileAccess.Read,
                                                                    FileShare.ReadWrite)))
                {
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                        {
                            continue;
                        }

                        var lineParts = dataLine.Split('\t');
                        if (lineParts.Length >= (int)eSICStatsColumns.StatMomentsArea + 1 &&
                            int.TryParse(lineParts[(int)eSICStatsColumns.FragScanNumber], out var fragScanNumber))
                        {
                            //  Note: the remaining values are stored as strings to prevent the number format from changing
                            var sicStatsEntry = new SICStatsData(fragScanNumber)
                            {
                                OptimalScanNumber = string.Copy(lineParts[(int)eSICStatsColumns.OptimalPeakApexScanNumber]),
                                PeakMaxIntensity = string.Copy(lineParts[(int)eSICStatsColumns.PeakMaxIntensity]),
                                PeakSignalToNoiseRatio = string.Copy(lineParts[(int)eSICStatsColumns.PeakSignalToNoiseRatio]),
                                FWHMInScans = string.Copy(lineParts[(int)eSICStatsColumns.FWHMInScans]),
                                PeakArea = string.Copy(lineParts[(int)eSICStatsColumns.PeakArea]),
                                ParentIonIntensity = string.Copy(lineParts[(int)eSICStatsColumns.ParentIonIntensity]),
                                ParentIonMZ = string.Copy(lineParts[(int)eSICStatsColumns.MZ]),
                                StatMomentsArea = string.Copy(lineParts[(int)eSICStatsColumns.StatMomentsArea])
                            };

                            dctSICStats.Add(fragScanNumber, sicStatsEntry);
                        }

                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadSICStatsFile", ex);
                return false;
            }

        }

        private bool ReadReporterIonStatsFile(
            string sourceDirectory,
            string reporterIonStatsFileName,
            IDictionary<int, ScanStatsData> dctScanStats,
            out string reporterIonHeaders)
        {
            var warningCount = 0;
            reporterIonHeaders = string.Empty;

            try
            {
                ShowMessage("  Reading: " + reporterIonStatsFileName);
                using (var reader = new StreamReader(new FileStream(Path.Combine(sourceDirectory, reporterIonStatsFileName), FileMode.Open,
                                                                    FileAccess.Read, FileShare.ReadWrite)))
                {
                    var linesRead = 0;
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                        {
                            continue;
                        }

                        linesRead++;
                        var lineParts = dataLine.Split('\t');
                        if (linesRead == 1)
                        {
                            //  This is the header line; we need to cache it
                            if (lineParts.Length >= (int)eReporterIonStatsColumns.ReporterIonIntensityMax + 1)
                            {
                                reporterIonHeaders = lineParts[(int)eReporterIonStatsColumns.CollisionMode];
                                reporterIonHeaders += '\t' + FlattenArray(lineParts, (int)eReporterIonStatsColumns.ReporterIonIntensityMax);
                            }
                            else
                            {
                                //  There aren't enough columns in the header line; this is unexpected
                                reporterIonHeaders = "Collision Mode" + '\t' + "AdditionalReporterIonColumns";
                            }

                        }

                        if (lineParts.Length < (int)eReporterIonStatsColumns.ReporterIonIntensityMax + 1)
                        {
                            continue;
                        }

                        if (!int.TryParse(lineParts[(int)eReporterIonStatsColumns.ScanNumber], out var scanNumber))
                        {
                            continue;
                        }

                        //  Look for scanNumber in scanNumbers
                        if (!dctScanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                        {
                            if (warningCount < 10)
                            {
                                ShowMessage("Warning: "
                                             + REPORTER_IONS_FILE_EXTENSION + " file refers to scan "
                                                                                + scanNumber +
                                                                                   ", but that scan was not in the _ScanStats.txt file");
                            }
                            else if (warningCount == 10)
                            {
                                ShowMessage("Warning: "
                                             + REPORTER_IONS_FILE_EXTENSION +
                                                " file has 10 or more scan numbers that are not defined in the _ScanStats.txt file");
                            }

                            warningCount++;
                        }
                        else if (scanStatsEntry.ScanNumber != scanNumber)
                        {
                            //  Scan number mismatch; this shouldn't happen
                            ShowMessage("Error: Scan number mismatch in ReadReporterIonStatsFile: "
                                         + scanStatsEntry.ScanNumber + " vs. " + scanNumber);
                        }
                        else
                        {
                            scanStatsEntry.CollisionMode = string.Copy(lineParts[(int)eReporterIonStatsColumns.CollisionMode]);
                            scanStatsEntry.ReporterIonData = FlattenArray(lineParts, (int)eReporterIonStatsColumns.ReporterIonIntensityMax);
                        }

                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadSICStatsFile", ex);
                return false;
            }
        }

        private void SetLocalErrorCode(eResultsProcessorErrorCodes eNewErrorCode, bool leaveExistingErrorCodeUnchanged = false)
        {
            if (leaveExistingErrorCodeUnchanged && mLocalErrorCode != eResultsProcessorErrorCodes.NoError)
            {
                //  An error code is already defined; do not change it
            }
            else
            {
                mLocalErrorCode = eNewErrorCode;
                if (eNewErrorCode == eResultsProcessorErrorCodes.NoError)
                {
                    if (ErrorCode == ProcessFilesErrorCodes.LocalizedError)
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError);
                    }

                }
                else
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError);
                }

            }

        }

        private KeyValuePair<string, string>[] SummarizeCollisionModes(
            FileSystemInfo inputFile,
            string baseFileName,
            string outputDirectoryPath,
            Dictionary<int, ScanStatsData> dctScanStats,
            out Dictionary<string, int> dctCollisionModeFileMap)
        {
            dctCollisionModeFileMap = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            var collisionModeTypeCount = 0;
            foreach (var scanStatsItem in dctScanStats.Values)
            {
                if (!dctCollisionModeFileMap.ContainsKey(scanStatsItem.CollisionMode))
                {
                    //  Store this collision mode in htCollisionModes; the value stored will be the index in collisionModes()
                    dctCollisionModeFileMap.Add(scanStatsItem.CollisionMode, collisionModeTypeCount);
                    collisionModeTypeCount++;
                }

            }

            if (dctCollisionModeFileMap.Count == 0 ||
                dctCollisionModeFileMap.Count == 1 && string.IsNullOrWhiteSpace(dctCollisionModeFileMap.First().Key))
            {
                //  Try to load the collision mode info from the input file
                //  MS-GF+ results report this in the FragMethod column
                dctCollisionModeFileMap.Clear();
                collisionModeTypeCount = 0;
                try
                {
                    using (var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                    {
                        var linesRead = 0;
                        var fragMethodColNumber = 0;
                        while (!reader.EndOfStream)
                        {
                            var dataLine = reader.ReadLine();
                            if (string.IsNullOrWhiteSpace(dataLine))
                            {
                                continue;
                            }

                            linesRead++;
                            var lineParts = dataLine.Split('\t');
                            if (linesRead == 1)
                            {
                                //  Header line; look for the FragMethod column
                                for (var colIndex = 0; colIndex < lineParts.Length; colIndex++)
                                {
                                    if (string.Equals(lineParts[colIndex], "FragMethod", StringComparison.OrdinalIgnoreCase))
                                    {
                                        fragMethodColNumber = colIndex + 1;
                                        break;
                                    }

                                }

                                if (fragMethodColNumber == 0)
                                {
                                    //  Fragmentation method column not found
                                    ShowWarning("Unable to determine the collision mode for results being merged. " +
                                                 "This is typically obtained from a MASIC _ReporterIons.txt file " +
                                                  "or from the FragMethod column in the MS-GF+ results file");
                                    break;
                                }

                                //  Also look for the scan number column and the protein column
                                FindScanNumColumn(inputFile, lineParts);
                                continue;
                            }

                            if (lineParts.Length < ScanNumberColumn ||
                                !int.TryParse(lineParts[ScanNumberColumn - 1], out var scanNumber))
                            {
                                continue;
                            }

                            if (lineParts.Length < fragMethodColNumber)
                            {
                                continue;
                            }

                            var collisionMode = lineParts[fragMethodColNumber - 1];
                            if (!dctCollisionModeFileMap.ContainsKey(collisionMode))
                            {
                                //  Store this collision mode in htCollisionModes; the value stored will be the index in collisionModes()
                                dctCollisionModeFileMap.Add(collisionMode, collisionModeTypeCount);
                                collisionModeTypeCount++;
                            }

                            if (!dctScanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                            {
                                scanStatsEntry = new ScanStatsData(scanNumber)
                                {
                                    CollisionMode = collisionMode
                                };
                                dctScanStats.Add(scanNumber, scanStatsEntry);
                            }
                            else
                            {
                                scanStatsEntry.CollisionMode = collisionMode;
                            }

                        }
                    }

                }
                catch (Exception ex)
                {
                    HandleException("Error extraction collision mode information from the input file", ex);
                    return new KeyValuePair<string, string>[collisionModeTypeCount - 1];
                }

            }

            if (collisionModeTypeCount == 0)
            {
                collisionModeTypeCount = 1;
            }

            var outputFilePaths = new KeyValuePair<string, string>[collisionModeTypeCount - 1];

            if (dctCollisionModeFileMap.Count == 0)
            {
                outputFilePaths[0] = new KeyValuePair<string, string>("na", Path.Combine(outputDirectoryPath, baseFileName + "_na" + RESULTS_SUFFIX));
            }
            else
            {
                foreach (var item in dctCollisionModeFileMap)
                {
                    var collisionMode = item.Key;
                    if (string.IsNullOrWhiteSpace(collisionMode))
                    {
                        collisionMode = "na";
                    }

                    var outputFilePath = Path.Combine(outputDirectoryPath, baseFileName + "_" + collisionMode + RESULTS_SUFFIX);
                    outputFilePaths[item.Value] = new KeyValuePair<string, string>(collisionMode, outputFilePath);
                }

            }

            return outputFilePaths;
        }
    }
}