using System.Collections.Generic;
using PHRPReader;

namespace MASICResultsMerger
{
    class DartIdData
    {
        public int Charge { get; set; }

        public string DataLine { get; }

        public string ElutionTime { get; set; }

        public string PeakWidthMinutes { get; set; }

        /// <summary>
        /// Peptide sequence, including prefix and suffix letters
        /// Optionally contains mod symbols
        /// </summary>
        public string Peptide { get; }

        /// <summary>
        /// Peptide sequence (possibly with mods), but without prefix and suffix residues
        /// </summary>
        public string PrimarySequence { get; }

        public List<string> Proteins { get; }

        public int ScanNumber { get; }

        public string SpecEValue { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public DartIdData() : this(string.Empty, 0, string.Empty, string.Empty, string.Empty)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="dataLine"></param>
        /// <param name="scanNumber"></param>
        /// <param name="peptide"></param>
        /// <param name="primarySequence"></param>
        /// <param name="proteinName"></param>
        public DartIdData(string dataLine, int scanNumber, string peptide, string primarySequence, string proteinName)
        {
            DataLine = dataLine;
            ScanNumber = scanNumber;
            Peptide = peptide;
            PrimarySequence = primarySequence;

            Proteins = new List<string> {
                proteinName
            };
        }
    }
}
