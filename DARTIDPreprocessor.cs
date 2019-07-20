using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PHRPReader;

namespace MASICResultsMerger
{
    /// <summary>
    /// Transforms a Dataset_PlusSICStats.txt data file from MASICResultsMerger into a format compatible with DART-ID
    /// Dart-ID Manuscript: PLoS Computational Biology. 2019 Jul 1;15(7):e1007082
    /// https://www.ncbi.nlm.nih.gov/pubmed/31260443
    /// </summary>
    class DartIdPreprocessor : PRISM.EventNotifier
    {

        private bool ColumnExists(IReadOnlyDictionary<Enum, int> msgfPlusColumns, clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns requiredColumn)
        {
            if (msgfPlusColumns.TryGetValue(requiredColumn, out var columnIndex))
            {
                if (columnIndex >= 0)
                {
                    return true;
                }
            }

            return false;
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
                var scanTimeColIndex = -1;

                var requiredColumns = new List<clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns>
                {
                    clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns.Peptide,
                    clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns.SpecProb_EValue,
                    clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns.Charge,
                    clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns.Protein
                };

                string datasetName;
                if (multiJobFile)
                {
                    datasetName = "TBD";
                }
                else
                {
                    //  Obtain the dataset name from the filename
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

                    //  ReSharper disable StringLiteralTypo
                    if (datasetName.EndsWith("_msgfplus", StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = datasetName.Substring(0, datasetName.Length - "_msgfplus".Length);
                    }
                    else if (datasetName.EndsWith("_msgfdb", StringComparison.OrdinalIgnoreCase))
                    {
                        datasetName = datasetName.Substring(0, datasetName.Length - "_msgfdb".Length);
                    }

                    //  ReSharper restore StringLiteralTypo
                }

                using (var reader = new StreamReader(new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                using (var writer = new StreamWriter(new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                            continue;

                        if (scanTimeColIndex < 0)
                        {
                            var success = ParseMergedFileHeaderLine(dataLine, ref scanTimeColIndex, msgfPlusColumns);
                            if (!success)
                            {
                                return false;
                            }

                            if (scanTimeColIndex < 0)
                            {
                                OnErrorEvent(string.Format("File {0} is missing column {1} on the header line", inputFile.Name,
                                                           MASICResultsMerger.SCAN_STATS_ELUTION_TIME_COLUMN));
                                return false;
                            }

                            //  Validate that the required columns exist
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

                        var columns = dataLine.Split('\t');
                        var peptide = clsPHRPReader.LookupColumnValue(columns, clsPHRPParserMSGFPlus.MSGFPlusSynFileColumns.Peptide,
                                                                      msgfPlusColumns, string.Empty);

                    }
                }

                throw new NotImplementedException();
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ConsolidatePSMs", ex);
                return false;
            }

        }

        private bool ParseMergedFileHeaderLine(string headerLine, ref int scanTimeColIndex, IDictionary<Enum, int> msgfPlusColumns)
        {
            try
            {
                var headerNames = headerLine.Split('\t').ToList();
                var columnNameToIndexMap = clsPHRPParserMSGFPlus.GetColumnMapFromHeaderLine(headerNames);
                foreach (var item in columnNameToIndexMap)
                {
                    msgfPlusColumns.Add(item.Key, item.Value);
                }

                for (var index = 0; index < headerNames.Count; index++)
                {
                    if (headerNames[index].Equals(MASICResultsMerger.SCAN_STATS_ELUTION_TIME_COLUMN, StringComparison.OrdinalIgnoreCase))
                    {
                        scanTimeColIndex = index;
                        break;
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
