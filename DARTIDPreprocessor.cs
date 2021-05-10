using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PHRPReader;
using PHRPReader.Data;
using PHRPReader.Reader;

namespace MASICResultsMerger
{
    /// <summary>
    /// Transforms a Dataset_PlusSICStats.txt data file from MASICResultsMerger into a format compatible with DART-ID
    /// Dart-ID Manuscript: PLoS Computational Biology. 2019 Jul 1;15(7):e1007082
    /// https://www.ncbi.nlm.nih.gov/pubmed/31260443
    /// </summary>
    internal class DartIdPreprocessor : PRISM.EventNotifier
    {
        // Ignore Spelling: tsv, fht, msgfdb

        private int mScanTimeColIndex;

        private int mPeakWidthMinutesColIndex;

        private bool ColumnExists(IReadOnlyDictionary<Enum, int> msgfPlusColumns, MSGFPlusSynFileColumns requiredColumn)
        {
            return msgfPlusColumns.TryGetValue(requiredColumn, out var columnIndex) && columnIndex >= 0;
        }

        public bool ConsolidatePSMs(string psmFilePath, bool multiJobFile)
        {
            try
            {
                var inputFile = new FileInfo(psmFilePath);
                var outputFileName = Path.GetFileNameWithoutExtension(inputFile.Name) + "_ForDartID.txt";
                string outputFilePath;

                if (inputFile.DirectoryName != null)
                {
                    outputFilePath = Path.Combine(inputFile.DirectoryName, outputFileName);
                }
                else
                {
                    outputFilePath = outputFileName;
                }

                var msgfPlusColumns = new SortedDictionary<Enum, int>();
                mScanTimeColIndex = -1;
                mPeakWidthMinutesColIndex = -1;

                var requiredColumns = new List<MSGFPlusSynFileColumns>
                {
                    MSGFPlusSynFileColumns.Peptide,
                    MSGFPlusSynFileColumns.SpecEValue,
                    MSGFPlusSynFileColumns.Charge,
                    MSGFPlusSynFileColumns.Protein
                };

                string datasetName;
                if (multiJobFile)
                {
                    datasetName = "TBD";
                    throw new NotImplementedException(
                        "ConsolidatePSMs needs to be updated to support an input file where Job or Dataset is the first column");
                }
                else
                {
                    // Obtain the dataset name from the filename
                    if (psmFilePath.EndsWith(MASICResultsMerger.RESULTS_SUFFIX, StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = Path.GetFileName(psmFilePath.Substring(0, psmFilePath.Length - MASICResultsMerger.RESULTS_SUFFIX.Length));
                    }
                    else
                    {
                        datasetName = Path.GetFileNameWithoutExtension(psmFilePath);
                    }

                    if (datasetName.EndsWith("_syn", StringComparison.OrdinalIgnoreCase) ||
                        datasetName.EndsWith("_fht", StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = datasetName.Substring(0, datasetName.Length - 4);
                    }

                    // ReSharper disable StringLiteralTypo
                    if (datasetName.EndsWith("_msgfplus", StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = datasetName.Substring(0, datasetName.Length - "_msgfplus".Length);
                    }
                    else if (datasetName.EndsWith("_msgfdb", StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = datasetName.Substring(0, datasetName.Length - "_msgfdb".Length);
                    }

                    // ReSharper restore StringLiteralTypo
                }

                var psmGroup = new DartIdData();

                using var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
                using var writer = new StreamWriter(new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));

                var headerLine = new List<string>
                {
                    "Dataset",
                    "Peptide",
                    "MSGFDB_SpecEValue",
                    "Charge",
                    "LeadingProtein",
                    "Proteins",
                    "ElutionTime",
                    "PeakWidthMinutes"
                };

                writer.WriteLine(string.Join("\t", headerLine));

                while (!reader.EndOfStream)
                {
                    var dataLine = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(dataLine))
                        continue;

                    if (mScanTimeColIndex < 0)
                    {
                        var success = ParseMergedFileHeaderLine(dataLine, msgfPlusColumns);
                        if (!success)
                        {
                            return false;
                        }

                        if (mScanTimeColIndex < 0)
                        {
                            OnErrorEvent(string.Format("File {0} is missing column {1} on the header line", inputFile.Name,
                                MASICResultsMerger.SCAN_STATS_ELUTION_TIME_COLUMN));
                            return false;
                        }

                        if (mPeakWidthMinutesColIndex < 0)
                        {
                            OnErrorEvent(string.Format("File {0} is missing column {1} on the header line", inputFile.Name,
                                MASICResultsMerger.PEAK_WIDTH_MINUTES_COLUMN));
                            return false;
                        }

                        // Validate that the required columns exist
                        foreach (var requiredColumn in requiredColumns)
                        {
                            if (!ColumnExists(msgfPlusColumns, requiredColumn))
                            {
                                OnErrorEvent(string.Format("File {0} is missing column {1} on the header line", inputFile.Name,
                                    requiredColumn.ToString()));
                                return false;
                            }
                        }

                        continue;
                    }

                    var dataColumns = dataLine.Split('\t');
                    var scanNumber = GetValueInt(dataColumns, msgfPlusColumns, MSGFPlusSynFileColumns.Scan);
                    var charge = GetValueInt(dataColumns, msgfPlusColumns, MSGFPlusSynFileColumns.Charge);
                    var peptide = GetValue(dataColumns, msgfPlusColumns, MSGFPlusSynFileColumns.Peptide);
                    var protein = GetValue(dataColumns, msgfPlusColumns, MSGFPlusSynFileColumns.Protein);

                    if (!PeptideCleavageStateCalculator.SplitPrefixAndSuffixFromSequence(peptide, out var primarySequence, out _, out _))
                    {
                        primarySequence = peptide;
                    }

                    if (scanNumber != psmGroup.ScanNumber ||
                        charge != psmGroup.Charge ||
                        !string.Equals(primarySequence, psmGroup.PrimarySequence))
                    {
                        StoreResult(writer, psmGroup, datasetName);

                        psmGroup = new DartIdData(dataLine, scanNumber, peptide, primarySequence, protein)
                        {
                            SpecEValue = GetValue(dataColumns, msgfPlusColumns, MSGFPlusSynFileColumns.SpecEValue),
                            Charge = charge,
                            ElutionTime = dataColumns[mScanTimeColIndex],
                            PeakWidthMinutes = dataColumns[mPeakWidthMinutesColIndex]
                        };
                    }
                    else
                    {
                        psmGroup.Proteins.Add(protein);
                    }
                }

                StoreResult(writer, psmGroup, datasetName);

                return true;
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ConsolidatePSMs", ex);
                return false;
            }
        }

        private void StoreResult(TextWriter writer, DartIdData groupData, string datasetName)
        {
            if (string.IsNullOrWhiteSpace(groupData.DataLine))
                return;

            var outputValues = new List<string>
            {
                datasetName,
                groupData.Peptide,
                groupData.SpecEValue,
                groupData.Charge.ToString(),
                groupData.Proteins.First(),
                string.Join(";", groupData.Proteins.Distinct()),
                groupData.ElutionTime,
                groupData.PeakWidthMinutes
            };

            writer.WriteLine(string.Join("\t", outputValues));
        }

        private string GetValue(string[] dataColumns, SortedDictionary<Enum, int> msgfPlusColumns, MSGFPlusSynFileColumns columnEnum)
        {
            return ReaderFactory.LookupColumnValue(dataColumns, columnEnum, msgfPlusColumns, string.Empty);
        }

        private int GetValueInt(string[] dataColumns, SortedDictionary<Enum, int> msgfPlusColumns, MSGFPlusSynFileColumns columnEnum)
        {
            var dataValue = GetValue(dataColumns, msgfPlusColumns, columnEnum);
            if (int.TryParse(dataValue, out var value))
                return value;

            return 0;
        }

        private bool ParseMergedFileHeaderLine(string headerLine, IDictionary<Enum, int> msgfPlusColumns)
        {
            try
            {
                var headerNames = headerLine.Split('\t').ToList();
                var columnNameToIndexMap = MSGFPlusSynFileReader.GetColumnMapFromHeaderLine(headerNames);
                foreach (var item in columnNameToIndexMap)
                {
                    msgfPlusColumns.Add(item.Key, item.Value);
                }

                for (var index = 0; index < headerNames.Count; index++)
                {
                    if (headerNames[index].Equals(MASICResultsMerger.SCAN_STATS_ELUTION_TIME_COLUMN, StringComparison.OrdinalIgnoreCase))
                    {
                        mScanTimeColIndex = index;
                    }
                    else if (headerNames[index].Equals(MASICResultsMerger.PEAK_WIDTH_MINUTES_COLUMN, StringComparison.OrdinalIgnoreCase))
                    {
                        mPeakWidthMinutesColIndex = index;
                    }
                    else if (headerNames[index].Equals("SpecEValue", StringComparison.OrdinalIgnoreCase))
                    {
                        // In .tsv files created by MzidToTsvConverter, the MSGFDB_SpecEValue column is named SpecEValue
                        if (!msgfPlusColumns.ContainsKey(MSGFPlusSynFileColumns.SpecEValue))
                        {
                            msgfPlusColumns.Add(MSGFPlusSynFileColumns.SpecEValue, index);
                        }
                    }
                    else if (headerNames[index].Equals("ScanNum", StringComparison.OrdinalIgnoreCase))
                    {
                        // In .tsv files created by MzidToTsvConverter, the Scan column is named ScanNum
                        if (!msgfPlusColumns.ContainsKey(MSGFPlusSynFileColumns.Scan))
                        {
                            msgfPlusColumns.Add(MSGFPlusSynFileColumns.Scan, index);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ParseHeaderLine", ex);
                return false;
            }
        }
    }
}
