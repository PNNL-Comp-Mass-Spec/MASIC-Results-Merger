using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using PHRPReader;
using PRISM;

namespace MASICResultsMerger
{
    /// <summary>
    /// This class merges the contents of a tab-delimited peptide hit results file from PHRP
    /// (for MS-GF+, MaxQuant, X!Tandem, etc.) with the corresponding MASIC results files,
    /// appending the relevant MASIC stats to each peptide hit result
    /// </summary>
    internal class MASICResultsMerger : PRISM.FileProcessor.ProcessFilesBase
    {
        // ReSharper disable once CommentTypo
        // Ignore Spelling: Frag, Mage, na, Num, SICstats

        /// <summary>
        /// Constructor
        /// </summary>
        public MASICResultsMerger()
        {
            mFileDate = "November 9, 2021";
            ProcessedDatasets = new List<ProcessedFileInfo>();

            InitializeLocalVariables();
        }

        #region "Constants and Enums"

        // ReSharper disable CommentTypo

        /// <summary>
        /// _SICstats filename suffix
        /// </summary>
        /// <remarks>MASIC uses lowercase "stats" in the SICstats filename</remarks>
        // ReSharper restore CommentTypo
        private const string SIC_STATS_FILE_EXTENSION = "_SICstats.txt";
        private const string SCAN_STATS_FILE_EXTENSION = "_ScanStats.txt";
        private const string REPORTER_IONS_FILE_EXTENSION = "_ReporterIons.txt";

        public const string RESULTS_SUFFIX = "_PlusSICStats.txt";
        public const int DEFAULT_SCAN_NUMBER_COLUMN = 2;

        public const string SCAN_STATS_ELUTION_TIME_COLUMN = "ElutionTime";

        public const string PEAK_WIDTH_MINUTES_COLUMN = "PeakWidthMinutes";

        /// <summary>
        /// Error codes specialized for this class
        /// </summary>
        public enum ResultsProcessorErrorCodes
        {
            NoError = 0,
            MissingMASICFiles = 1,
            MissingMageFiles = 2,
            UnspecifiedError = -1,
        }

        #endregion

        #region "Class wide Variables"

        private ResultsProcessorErrorCodes mLocalErrorCode;

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
        /// For the input file, defines which column tracks scan number; the first column is column 1 (not zero)
        /// </summary>
        public int ScanNumberColumn { get; set; }

        /// <summary>
        /// When true, show additional debug messages
        /// </summary>
        public bool TraceMode { get; set; }

        #endregion

        private double ComputePeakWidthMinutes(IReadOnlyDictionary<int, ScanStatsData> scanStats, string peakScanStart, string peakScanEnd)
        {
            if (!int.TryParse(peakScanStart, out var startScan))
                return 0;

            if (!int.TryParse(peakScanEnd, out var endScan))
                return 0;

            if (!scanStats.TryGetValue(startScan, out var startScanStats))
                return 0;

            if (!scanStats.TryGetValue(endScan, out var endScanStats))
                return 0;

            if (!double.TryParse(startScanStats.ElutionTime, out var startTimeMinutes))
                return 0;

            if (!double.TryParse(endScanStats.ElutionTime, out var endTimeMinutes))
                return 0;

            return endTimeMinutes - startTimeMinutes;
        }

        private bool FindMASICFiles(
            string masicResultsDirectory,
            DatasetInfo datasetInfo,
            MASICFileInfo masicFiles,
            string masicFileSearchInfo,
            int job)
        {
            var triedDatasetID = false;
            var success = false;

            try
            {
                Console.WriteLine();
                ShowMessage("Looking for MASIC data files that correspond to " + datasetInfo.DatasetName);

                var datasetName = datasetInfo.DatasetName;

                // Use a loop to try various possible dataset names
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

                    if (TraceMode)
                    {
                        OnDebugEvent("These expected files were not found:\n   {0}\n   {1}", scanStatsFile.Name, sicStatsFile.Name);
                        Console.WriteLine();
                    }

                    // Find the last underscore in datasetName, then remove it and any text after it
                    var charIndex = datasetName.LastIndexOf('_');
                    if (charIndex > 0)
                    {
                        datasetName = datasetName.Substring(0, charIndex);

                        if (TraceMode)
                        {
                            ShowMessage("Now looking for " + datasetName);
                        }
                    }
                    else if (!triedDatasetID && datasetInfo.DatasetID > 0)
                    {
                        datasetName = datasetInfo.DatasetID + "_" + datasetInfo.DatasetName;
                        triedDatasetID = true;

                        if (TraceMode)
                        {
                            ShowMessage("Now looking for " + datasetName);
                        }
                    }
                    else
                    {
                        // No more underscores; we're unable to determine the dataset name
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
                // ReSharper disable once StringLiteralTypo
                ShowMessage("  Note: The MASIC SICstats file was found, but the ScanStats file does not exist " + masicFileSearchInfo);
            }
            else if (string.IsNullOrWhiteSpace(masicFiles.SICStatsFileName))
            {
                // ReSharper disable once StringLiteralTypo
                ShowMessage("  Note: The MASIC ScanStats file was found, but the SICstats file does not exist " + masicFileSearchInfo);
            }

            return true;
        }

        private void FindScanNumColumn(FileSystemInfo inputFile, IList<string> lineParts)
        {
            // Check for a column named "ScanNum" or "ScanNumber" or "Scan Num" or "Scan Number"
            // If found, override ScanNumberColumn
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
            return new()
            {
                SCAN_STATS_ELUTION_TIME_COLUMN,
                "ScanType",
                "TotalIonIntensity",
                "BasePeakIntensity",
                "BasePeakMZ"
            };
        }

        private List<string> GetSICStatsHeaders()
        {
            return new()
            {
                "Optimal_Scan_Number",
                "PeakMaxIntensity",
                "PeakSignalToNoiseRatio",
                "FWHMInScans",
                "PeakArea",
                "ParentIonIntensity",
                "ParentIonMZ",
                "StatMomentsArea",
                "PeakScanStart",
                "PeakScanEnd",
                PEAK_WIDTH_MINUTES_COLUMN
            };
        }

        private string FlattenList(IEnumerable<string> lstData)
        {
            return string.Join("\t", lstData);
        }

        /// <summary>
        /// Get the current error state, if any
        /// </summary>
        /// <returns>Returns an empty string if no error</returns>
        public override string GetErrorMessage()
        {
            if (ErrorCode is ProcessFilesErrorCodes.LocalizedError or ProcessFilesErrorCodes.NoError)
            {
                return mLocalErrorCode switch
                {
                    ResultsProcessorErrorCodes.NoError => string.Empty,
                    ResultsProcessorErrorCodes.UnspecifiedError => "Unspecified localized error",
                    ResultsProcessorErrorCodes.MissingMASICFiles => "Missing MASIC Files",
                    ResultsProcessorErrorCodes.MissingMageFiles => "Missing Mage Extractor Files",
                    _ => "Unknown error state",
                };
            }

            return GetBaseClassErrorMessage();
        }

        private void InitializeLocalVariables()
        {
            mMASICResultsDirectoryPath = string.Empty;
            ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN;
            SeparateByCollisionMode = false;
            mLocalErrorCode = ResultsProcessorErrorCodes.NoError;
        }

        private bool MergePeptideHitAndMASICFiles(
            FileSystemInfo inputFile,
            string outputDirectoryPath,
            Dictionary<int, ScanStatsData> scanStats,
            IReadOnlyDictionary<int, SICStatsData> sicStats,
            string reporterIonHeaders)
        {
            StreamWriter[] writers;
            int[] linesWritten;

            int outputFileCount;

            // The Key is the collision mode and the value is the path
            KeyValuePair<string, string>[] outputFilePaths;
            string baseFileName;

            var blankAdditionalScanStatsColumns = string.Empty;
            var blankAdditionalSICColumns = string.Empty;
            var blankAdditionalReporterIonColumns = string.Empty;

            Dictionary<string, int> collisionModeFileMap;

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

                // Define the output file path
                outputDirectoryPath ??= string.Empty;

                if (SeparateByCollisionMode)
                {
                    outputFilePaths = SummarizeCollisionModes(inputFile, baseFileName, outputDirectoryPath, scanStats, out collisionModeFileMap);
                    outputFileCount = outputFilePaths.Length;
                    if (outputFileCount < 1)
                    {
                        return false;
                    }
                }
                else
                {
                    collisionModeFileMap = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
                    outputFileCount = 1;
                    outputFilePaths = new KeyValuePair<string, string>[outputFileCount];
                    outputFilePaths[0] = new KeyValuePair<string, string>(string.Empty, Path.Combine(outputDirectoryPath, baseFileName + RESULTS_SUFFIX));
                }

                writers = new StreamWriter[outputFileCount];
                linesWritten = new int[outputFileCount];

                for (var index = 0; index < outputFileCount; index++)
                {
                    writers[index] = new StreamWriter(
                        new FileStream(outputFilePaths[index].Value, FileMode.Create, FileAccess.Write, FileShare.Read));
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
                    // Assume the scan number is in the second column
                    ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN;
                }

                // Read from reader and write out to the file(s) in writers[]
                using (var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    var linesRead = 0;
                    var writeReporterIonStats = false;
                    var writeSICStats = (sicStats.Count > 0);

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

                            // Write out an updated header line
                            if (lineParts.Count >= ScanNumberColumn && int.TryParse(lineParts[ScanNumberColumn - 1], out _))
                            {
                                // The input file doesn't have a header line; we will add one, using generic column names for the data in the input file
                                var genericHeaders = new List<string>();
                                for (var index = 0; index < lineParts.Count; index++)
                                {
                                    genericHeaders.Add("Column" + index.ToString("00"));
                                }

                                headerLine = FlattenList(genericHeaders);
                            }
                            else
                            {
                                // The input file does have a text-based header
                                headerLine = dataLine;
                                FindScanNumColumn(inputFile, lineParts);

                                // Clear splitLine so that this line gets skipped
                                lineParts.Clear();
                            }

                            var scanStatsHeaders = GetScanStatsHeaders();
                            var sicStatsHeaders = GetSICStatsHeaders();

                            if (!writeSICStats)
                            {
                                sicStatsHeaders.Clear();
                            }

                            // Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
                            blankAdditionalScanStatsColumns = new string('\t', scanStatsHeaders.Count - 1);
                            if (writeSICStats)
                            {
                                blankAdditionalSICColumns = new string('\t', sicStatsHeaders.Count);
                            }

                            // Initialize blankAdditionalReporterIonColumns
                            if (reporterIonHeaders.Length > 0)
                            {
                                blankAdditionalReporterIonColumns = new string('\t', reporterIonHeaders.Split('\t').ToList().Count - 1);
                            }

                            // Initialize the AddOn header columns
                            var addonHeaders = FlattenList(scanStatsHeaders);
                            if (writeSICStats)
                            {
                                addonHeaders += '\t' + FlattenList(sicStatsHeaders);
                            }

                            if (reporterIonHeaders.Length > 0)
                            {
                                // Append the reporter ion stats columns
                                addonHeaders += '\t' + reporterIonHeaders;
                                writeReporterIonStats = true;
                            }

                            // Write out the headers
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

                        // Look for scanNumber in scanStats
                        var addonColumns = new List<string>();
                        if (!scanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                        {
                            // Match not found; use the blank columns in blankAdditionalScanStatsColumns
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
                            if (!sicStats.TryGetValue(scanNumber, out var sicStatsEntry))
                            {
                                // Match not found; use the blank columns in blankAdditionalSICColumns
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
                                addonColumns.Add(sicStatsEntry.PeakScanStart);
                                addonColumns.Add(sicStatsEntry.PeakScanEnd);

                                var peakWidthMinutes = ComputePeakWidthMinutes(scanStats, sicStatsEntry.PeakScanStart, sicStatsEntry.PeakScanEnd);
                                addonColumns.Add(StringUtilities.DblToString(peakWidthMinutes, 4));
                            }
                        }

                        if (writeReporterIonStats)
                        {
                            if (scanStatsEntry == null || string.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode))
                            {
                                // Collision mode is not defined; append blank columns
                                addonColumns.Add(blankAdditionalReporterIonColumns);
                            }
                            else
                            {
                                // Collision mode is defined
                                addonColumns.Add(scanStatsEntry.CollisionMode);
                                addonColumns.Add(scanStatsEntry.ReporterIonData);
                                collisionModeCurrentScan = scanStatsEntry.CollisionMode;
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
                                collisionModeCurrentScan = scanStatsEntry.CollisionMode;
                            }
                        }

                        var outFileIndex = 0;
                        if (SeparateByCollisionMode && outputFileCount > 1)
                        {
                            if (collisionModeCurrentScan != null)
                            {
                                // Determine the correct output file
                                if (!collisionModeFileMap.TryGetValue(collisionModeCurrentScan, out outFileIndex))
                                {
                                    outFileIndex = 0;
                                }
                            }
                        }

                        writers[outFileIndex].WriteLine(dataLine + '\t' + FlattenList(addonColumns));
                        linesWritten[outFileIndex]++;
                    }
                }

                // Close the output files
                if (writers != null)
                {
                    for (var index = 0; index < outputFileCount; index++)
                    {
                        writers[index]?.Close();
                    }
                }

                if (CreateDartIdInputFile)
                {
                    var preprocessor = new DartIdPreprocessor();

                    foreach (var item in outputFilePaths.ToList())
                    {
                        preprocessor.ConsolidatePSMs(item.Value, false);
                    }
                }

                // See if any of the files had no data written to them
                // If there are, then delete the empty output file
                // However, retain at least one output file
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
                        // All output files are empty
                        // Pretend the first output file actually contains data
                        linesWritten[0] = 1;
                    }

                    for (var index = 0; index < outputFileCount; index++)
                    {
                        // Wait 250 msec before continuing
                        Thread.Sleep(250);
                        if (linesWritten[index] == 0)
                        {
                            var fileName = Path.GetFileName(outputFilePaths[index].Value);

                            try
                            {
                                ShowMessage("Deleting empty output file: " + Environment.NewLine +
                                            " --> " + fileName);
                                File.Delete(outputFilePaths[index].Value);
                            }
                            catch (Exception ex)
                            {
                                ConsoleMsgUtils.ShowDebug("Unable to delete empty output file named '{0}': {1}", fileName, ex.Message);
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
                    // Nothing to merge
                    ShowMessage("Only one dataset has been processed by the MASICResultsMerger; nothing to merge");
                    return true;
                }

                // Determine the base filename and collision modes used
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

                    // Find the characters common to all of the processed datasets
                    var candidateName = processedDataset.BaseName;
                    if (string.IsNullOrEmpty(baseFileName))
                    {
                        baseFileName = candidateName;
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
                            // Possibly backtrack to the previous underscore
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

                // Open the output files
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

                // Create the DatasetMap file
                using (var writer = new StreamWriter(
                    new FileStream(Path.Combine(mOutputDirectoryPath, baseFileName + "_DatasetMap.txt"), FileMode.Create, FileAccess.Write)))
                {
                    writer.WriteLine("DatasetID" + '\t' + "DatasetName");
                    foreach (var datasetMapping in datasetNameIdMap)
                    {
                        writer.WriteLine(datasetMapping.Value + '\t' + datasetMapping.Key);
                    }
                }

                // Merge the files
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

                        using var reader = new StreamReader(new FileStream(sourceFile.Value, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

                        var linesRead = 0;
                        while (!reader.EndOfStream)
                        {
                            var dataLine = reader.ReadLine();
                            linesRead++;
                            if (linesRead == 1)
                            {
                                if (outputFileHeaderWritten[collisionMode])
                                {
                                    // skip this line
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
            string masicResultsDirectory;
            if (resetErrorCode)
            {
                SetLocalErrorCode(ResultsProcessorErrorCodes.NoError);
            }

            if (string.IsNullOrEmpty(inputFilePath))
            {
                ShowMessage("Input file name is empty");
                SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath);
                return false;
            }

            // Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
            if (!CleanupFilePaths(ref inputFilePath, ref outputDirectoryPath))
            {
                SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError);
                return false;
            }

            var inputFile = new FileInfo(inputFilePath);
            if (string.IsNullOrWhiteSpace(mMASICResultsDirectoryPath))
            {
                masicResultsDirectory = inputFile.DirectoryName;
            }
            else
            {
                masicResultsDirectory = mMASICResultsDirectoryPath;
            }

            ProcessedDatasets.Clear();

            if (MageResults)
            {
                return ProcessMageExtractorFile(inputFile, masicResultsDirectory);
            }

            return ProcessSingleJobFile(inputFile, masicResultsDirectory);
        }

        private bool ProcessMageExtractorFile(FileInfo inputFile, string masicResultsDirectory)
        {
            var scanStats = new Dictionary<int, ScanStatsData>();
            var sicStats = new Dictionary<int, SICStatsData>();
            try
            {
                // Read the Mage Metadata file
                FileInfo metadataFile;
                var metadataFileName = Path.GetFileNameWithoutExtension(inputFile.Name) + "_metadata.txt";

                if (inputFile.DirectoryName != null)
                {
                    metadataFile = new FileInfo(Path.Combine(inputFile.DirectoryName, metadataFileName));
                }
                else
                {
                    metadataFile = new FileInfo(metadataFileName);
                }

                if (!metadataFile.Exists)
                {
                    ShowErrorMessage("Error: Mage Metadata File not found: " + metadataFile.FullName);
                    SetLocalErrorCode(ResultsProcessorErrorCodes.MissingMageFiles);
                    return false;
                }

                // Keys in this dictionary are the job, values are the DatasetID and DatasetName
                var jobToDatasetMap = ReadMageMetadataFile(metadataFile.FullName);
                if (jobToDatasetMap == null || jobToDatasetMap.Count == 0)
                {
                    ShowErrorMessage("Error: ReadMageMetadataFile returned an empty job mapping");
                    return false;
                }

                string headerLine;
                int jobColumnIndex;

                // Open the Mage Extractor data file so that we can validate and cache the header row
                using (var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    headerLine = reader.ReadLine();
                    var invalidFileMessage = string.Format(
                        "Input file is not a valid Mage Extractor results file; it must contain a \"Job\" column: " +
                        inputFile.FullName);

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

                // Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
                var blankAdditionalScanStatsColumns = new string('\t', scanStatsHeaders.Count - 1);
                var blankAdditionalSICColumns = new string('\t', sicStatsHeaders.Count);

                var outputFileName = Path.GetFileNameWithoutExtension(inputFile.Name) + RESULTS_SUFFIX;
                var outputFilePath = Path.Combine(mOutputDirectoryPath, outputFileName);

                var jobsSuccessfullyMerged = 0;

                // Initialize the output file
                using (var writer = new StreamWriter(new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read)))
                {
                    // Open the Mage Extractor data file and read the data for each job
                    var phrpReader = new ReaderFactory(inputFile.FullName, PeptideHitResultTypes.Unknown, false, false, false)
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

                        // Parse out the job from the current line
                        var lstColumns = psm.DataLineText.Split('\t').ToList();

                        if (!int.TryParse(lstColumns[jobColumnIndex], out var job))
                        {
                            ShowMessage("Warning: Job column does not contain a job number; skipping this entry: " + psm.DataLineText);
                            continue;
                        }

                        if (job != lastJob)
                        {
                            // New job; read and cache the MASIC data
                            masicDataLoaded = false;

                            if (!jobToDatasetMap.TryGetValue(job, out var datasetInfo))
                            {
                                ShowErrorMessage("Error: Job " + job + " was not defined in the Metadata file; unable to determine the dataset");
                            }
                            else
                            {
                                // Look for the corresponding MASIC files in the input directory
                                var masicFiles = new MASICFileInfo();
                                var datasetNameAndDirectory = "for dataset " + datasetInfo.DatasetName + " in " + masicResultsDirectory;
                                var success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, datasetNameAndDirectory, job);
                                if (success)
                                {
                                    // Read and cache the MASIC data
                                    scanStats = new Dictionary<int, ScanStatsData>();
                                    sicStats = new Dictionary<int, SICStatsData>();

                                    masicDataLoaded = ReadMASICData(masicResultsDirectory, masicFiles, scanStats, sicStats, out reporterIonHeaders);
                                    if (masicDataLoaded)
                                    {
                                        jobsSuccessfullyMerged++;
                                        if (jobsSuccessfullyMerged == 1)
                                        {
                                            // Initialize blankAdditionalReporterIonColumns
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
                                    // The merged data file will have reporter ion columns
                                    writeReporterIonStats = true;
                                    writer.WriteLine(headerLine + '\t' + addonHeaderColumns + '\t' + reporterIonHeaders);
                                }
                                else
                                {
                                    writer.WriteLine(headerLine + '\t' + addonHeaderColumns);
                                }

                                headerLineWritten = true;
                            }

                            // Look for scanNumber in scanStats
                            var addonColumns = new List<string>();
                            if (!scanStats.TryGetValue(psm.ScanNumber, out var scanStatsEntry))
                            {
                                // Match not found; use the blank columns in blankAdditionalColumns
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

                            if (!sicStats.TryGetValue(psm.ScanNumber, out var sicStatsEntry))
                            {
                                // Match not found; use the blank columns in blankAdditionalSICColumns
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
                                addonColumns.Add(sicStatsEntry.PeakScanStart);
                                addonColumns.Add(sicStatsEntry.PeakScanEnd);

                                var peakWidthMinutes = ComputePeakWidthMinutes(scanStats, sicStatsEntry.PeakScanStart, sicStatsEntry.PeakScanEnd);
                                addonColumns.Add(StringUtilities.DblToString(peakWidthMinutes, 4));
                            }

                            if (writeReporterIonStats)
                            {
                                if (scanStatsEntry == null || string.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode))
                                {
                                    // Collision mode is not defined; append blank columns
                                    addonColumns.Add(blankAdditionalReporterIonColumns);
                                }
                                else
                                {
                                    // Collision mode is defined
                                    addonColumns.Add(scanStatsEntry.CollisionMode);
                                    addonColumns.Add(scanStatsEntry.ReporterIonData);
                                }
                            }

                            writer.WriteLine(psm.DataLineText + '\t' + FlattenList(addonColumns));
                        }
                        else
                        {
                            if (!headerLineWritten)
                            {
                                writer.WriteLine(headerLine);
                                headerLineWritten = true;
                            }

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

                        UpdateProgress("Loading data from " + inputFile.Name, phrpReader.PercentComplete);
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
                    preprocessor.ConsolidatePSMs(outputFilePath, true);
                }

                return jobsSuccessfullyMerged > 0;
            }
            catch (Exception ex)
            {
                HandleException("Error in ProcessMageExtractorFile", ex);
                return false;
            }
        }

        private bool ProcessSingleJobFile(FileSystemInfo inputFile, string masicResultsDirectory)
        {
            try
            {
                var datasetName = Path.GetFileNameWithoutExtension(inputFile.FullName);
                var datasetInfo = new DatasetInfo(datasetName, 0);

                // Note that FindMASICFiles will first try the full filename, and if it doesn't find a match,
                // it will start removing text from the end of the filename by looking for underscores
                // Look for the corresponding MASIC files in the input directory
                var masicFiles = new MASICFileInfo();
                var masicFileSearchInfo = "in " + masicResultsDirectory;

                var success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, masicFileSearchInfo, 0);

                if (!success)
                {
                    SetLocalErrorCode(ResultsProcessorErrorCodes.MissingMASICFiles);
                    return false;
                }

                // Read and cache the MASIC data
                var scanStats = new Dictionary<int, ScanStatsData>();
                var sicStats = new Dictionary<int, SICStatsData>();

                success = ReadMASICData(masicResultsDirectory, masicFiles, scanStats, sicStats, out var reporterIonHeaders);

                if (success)
                {
                    // Merge the MASIC data with the input file
                    success = MergePeptideHitAndMASICFiles(inputFile, mOutputDirectoryPath, scanStats, sicStats, reporterIonHeaders);
                }

                if (success)
                {
                    ShowMessage(string.Empty, false);
                }
                else
                {
                    SetLocalErrorCode(ResultsProcessorErrorCodes.UnspecifiedError);
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
            IDictionary<int, ScanStatsData> scanStats,
            IDictionary<int, SICStatsData> sicStats,
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
                    scanStatsRead = ReadScanStatsFile(sourceDirectory, masicFiles.ScanStatsFileName, scanStats);
                }

                if (string.IsNullOrWhiteSpace(masicFiles.SICStatsFileName))
                {
                    sicStatsRead = false;
                }
                else
                {
                    sicStatsRead = ReadSICStatsFile(sourceDirectory, masicFiles.SICStatsFileName, sicStats);
                }

                if (string.IsNullOrWhiteSpace(masicFiles.ReporterIonsFileName))
                {
                    reporterIonHeaders = string.Empty;
                }
                else
                {
                    ReadReporterIonStatsFile(sourceDirectory, masicFiles.ReporterIonsFileName, scanStats, out var reporterIonHeadersFromReader);

                    // Combine "Collision Mode" with the reporter ion headers obtained from ReadReporterIonStatsFile
                    reporterIonHeaders = "Collision Mode\t" + reporterIonHeadersFromReader;
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

        private bool ReadScanStatsFile(string sourceDirectory, string scanStatsFileName, IDictionary<int, ScanStatsData> scanStats)
        {
            try
            {
                // Initialize scanStats
                scanStats.Clear();
                ShowMessage("  Reading: " + scanStatsFileName);

                var reader = new PHRPReader.Reader.ScanStatsReader();
                RegisterEvents(reader);

                var scanStatsData = reader.ReadScanStatsData(Path.Combine(sourceDirectory, scanStatsFileName));

                foreach (var item in scanStatsData)
                {
                    var scanNumber = item.Value.ScanNumber;

                    // Note: the remaining values are stored as strings to prevent the number format from changing
                    var scanStatsEntry = new ScanStatsData(scanNumber)
                    {
                        ElutionTime = item.Value.ScanTimeMinutesText,
                        ScanType = item.Value.ScanType.ToString(),
                        TotalIonIntensity = item.Value.TotalIonIntensityText,
                        BasePeakIntensity = item.Value.BasePeakIntensityText,
                        BasePeakMZ = item.Value.BasePeakMzText,
                        CollisionMode = string.Empty,
                        ReporterIonData = string.Empty
                    };

                    scanStats.Add(scanNumber, scanStatsEntry);
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
            var jobToDatasetMap = new Dictionary<int, DatasetInfo>();
            var headersParsed = false;
            var jobIndex = -1;
            var datasetIndex = -1;
            var datasetIDIndex = -1;

            try
            {
                using var reader = new StreamReader(new FileStream(metadataFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

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
                        // Look for the Job and Dataset columns
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
                                jobToDatasetMap.Add(jobNumber, datasetInfo);
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
            catch (Exception ex)
            {
                HandleException("Error in ReadMageMetadataFile", ex);
                return null;
            }

            return jobToDatasetMap;
        }

        private bool ReadSICStatsFile(string sourceDirectory, string sicStatsFileName, IDictionary<int, SICStatsData> sicStats)
        {
            try
            {
                // Initialize sicStats
                sicStats.Clear();
                ShowMessage("  Reading: " + sicStatsFileName);

                var reader = new PHRPReader.Reader.SICStatsReader();
                RegisterEvents(reader);

                var sicStatsData = reader.ReadSICStatsData(Path.Combine(sourceDirectory, sicStatsFileName));

                foreach (var item in sicStatsData)
                {
                    var fragScanNumber = item.Value.FragScanNumber;

                    var sicStatsEntry = new SICStatsData(fragScanNumber)
                    {
                        OptimalScanNumber = item.Value.OptimalPeakApexScanNumber.ToString(),
                        PeakMaxIntensity = item.Value.PeakMaxIntensityText,
                        PeakSignalToNoiseRatio = item.Value.PeakSignalToNoiseRatioText,
                        FWHMInScans = item.Value.FWHMInScans.ToString(),
                        PeakScanStart = item.Value.PeakScanStart.ToString(),
                        PeakScanEnd = item.Value.PeakScanEnd.ToString(),
                        PeakArea = item.Value.PeakAreaText,
                        ParentIonIntensity = item.Value.ParentIonIntensityText,
                        ParentIonMZ = item.Value.MzText,
                        StatMomentsArea = item.Value.StatMomentsAreaText
                    };

                    sicStats.Add(fragScanNumber, sicStatsEntry);
                }

                return true;
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadSICStatsFile", ex);
                return false;
            }
        }

        private void ReadReporterIonStatsFile(
            string sourceDirectory,
            string reporterIonStatsFileName,
            IDictionary<int, ScanStatsData> scanStats,
            out string reporterIonHeaders)
        {
            var warningCount = 0;
            reporterIonHeaders = string.Empty;

            try
            {
                ShowMessage("  Reading: " + reporterIonStatsFileName);

                var reader = new PHRPReader.Reader.ReporterIonsFileReader();
                RegisterEvents(reader);

                var reporterIonData = reader.ReadReporterIonData(Path.Combine(sourceDirectory, reporterIonStatsFileName));
                reporterIonHeaders = FlattenList(reader.ReporterIonHeaderNames);

                foreach (var item in reporterIonData)
                {
                    var scanNumber = item.Value.ScanNumber;

                    // Look for scanNumber in scanNumbers
                    if (!scanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                    {
                        if (warningCount < 10)
                        {
                            ShowMessage(string.Format(
                                "Warning: the {0} file refers to scan {1}, but that scan was not in the _ScanStats.txt file",
                                REPORTER_IONS_FILE_EXTENSION, scanNumber));
                        }
                        else if (warningCount == 10)
                        {
                            ShowMessage(string.Format(
                                "Warning: the {0} file has 10 or more scan numbers that are not defined in the _ScanStats.txt file",
                                REPORTER_IONS_FILE_EXTENSION));
                        }

                        warningCount++;
                    }
                    else if (scanStatsEntry.ScanNumber != scanNumber)
                    {
                        // Scan number mismatch; this shouldn't happen
                        ShowMessage(string.Format(
                            "Error: Scan number mismatch in ReadReporterIonStatsFile: {0} vs. {1}",
                            scanStatsEntry.ScanNumber, scanNumber));
                    }
                    else
                    {
                        scanStatsEntry.CollisionMode = item.Value.CollisionMode;
                        scanStatsEntry.ReporterIonData = item.Value.ReporterIonColumnData;
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException("Error in ReadReporterIonStatsFile", ex);
            }
        }

        private void SetLocalErrorCode(ResultsProcessorErrorCodes eNewErrorCode, bool leaveExistingErrorCodeUnchanged = false)
        {
            if (leaveExistingErrorCodeUnchanged && mLocalErrorCode != ResultsProcessorErrorCodes.NoError)
            {
                // An error code is already defined; do not change it
            }
            else
            {
                mLocalErrorCode = eNewErrorCode;
                if (eNewErrorCode == ResultsProcessorErrorCodes.NoError)
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
            Dictionary<int, ScanStatsData> scanStats,
            out Dictionary<string, int> collisionModeFileMap)
        {
            collisionModeFileMap = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            var collisionModeTypeCount = 0;

            foreach (var scanStatsItem in scanStats.Values)
            {
                if (!collisionModeFileMap.ContainsKey(scanStatsItem.CollisionMode))
                {
                    // Store this collision mode in collisionModeFileMap; the value stored will be the index in collisionModes()
                    collisionModeFileMap.Add(scanStatsItem.CollisionMode, collisionModeTypeCount);
                    collisionModeTypeCount++;
                }
            }

            if (collisionModeFileMap.Count == 0 ||
                collisionModeFileMap.Count == 1 && string.IsNullOrWhiteSpace(collisionModeFileMap.First().Key))
            {
                // Try to load the collision mode info from the input file
                // MS-GF+ results report this in the FragMethod column
                collisionModeFileMap.Clear();
                collisionModeTypeCount = 0;
                try
                {
                    using var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

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
                            // Header line; look for the FragMethod column
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
                                // Fragmentation method column not found
                                ShowWarning("Unable to determine the collision mode for results being merged. " +
                                            "This is typically obtained from a MASIC _ReporterIons.txt file " +
                                            "or from the FragMethod column in the MS-GF+ results file");
                                break;
                            }

                            // Also look for the scan number column and the protein column
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

                        if (!collisionModeFileMap.ContainsKey(collisionMode))
                        {
                            // Store this collision mode in collisionModeFileMap; the value stored will be the index in collisionModes()
                            collisionModeFileMap.Add(collisionMode, collisionModeTypeCount);
                            collisionModeTypeCount++;
                        }

                        if (!scanStats.TryGetValue(scanNumber, out var scanStatsEntry))
                        {
                            scanStatsEntry = new ScanStatsData(scanNumber)
                            {
                                CollisionMode = collisionMode
                            };
                            scanStats.Add(scanNumber, scanStatsEntry);
                        }
                        else
                        {
                            scanStatsEntry.CollisionMode = collisionMode;
                        }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Error extraction collision mode information from the input file", ex);
                    return new KeyValuePair<string, string>[collisionModeTypeCount];
                }
            }

            if (collisionModeTypeCount == 0)
            {
                collisionModeTypeCount = 1;
            }

            var outputFilePaths = new KeyValuePair<string, string>[collisionModeTypeCount];

            if (collisionModeFileMap.Count == 0)
            {
                outputFilePaths[0] = new KeyValuePair<string, string>("na", Path.Combine(outputDirectoryPath, baseFileName + "_na" + RESULTS_SUFFIX));
            }
            else
            {
                foreach (var item in collisionModeFileMap)
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