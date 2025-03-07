namespace MASICResultsMerger
{
    internal class SICStatsData
    {
        // Ignore Spelling: frag

        public int FragScanNumber { get; }
        public string OptimalScanNumber { get; set; }
        public string PeakMaxIntensity { get; set; }
        public string PeakSignalToNoiseRatio { get; set; }
        public string FWHMInScans { get; set; }
        public string PeakScanStart { get; set; }
        public string PeakScanEnd { get; set; }
        public string PeakArea { get; set; }
        public string ParentIonIntensity { get; set; }
        public string ParentIonMZ { get; set; }
        public string StatMomentsArea { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fragScanNumber">Fragmentation scan number</param>
        public SICStatsData(int fragScanNumber)
        {
            FragScanNumber = fragScanNumber;
        }

        /// <summary>
        /// Compare the fragmentation scan numbers between two SIC stats items
        /// </summary>
        /// <param name="other">SIC stats instance to compare</param>
        /// <returns>Comparison result</returns>
        public int CompareTo(SICStatsData other)
        {
            if (FragScanNumber < other.FragScanNumber)
            {
                return -1;
            }

            // ReSharper disable once ConvertIfStatementToReturnStatement
            if (FragScanNumber > other.FragScanNumber)
            {
                return 1;
            }

            return 0;
        }
    }
}
