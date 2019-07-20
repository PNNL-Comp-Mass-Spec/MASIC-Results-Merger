namespace MASICResultsMerger
{
    class SICStatsData
    {

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
        /// <param name="fragScanNumber"></param>
        public SICStatsData(int fragScanNumber)
        {
            FragScanNumber = fragScanNumber;
        }

        public int CompareTo(SICStatsData other)
        {
            if (FragScanNumber < other.FragScanNumber)
            {
                return -1;
            }

            if (FragScanNumber > other.FragScanNumber)
            {
                return 1;
            }

            return 0;
        }
    }

}
