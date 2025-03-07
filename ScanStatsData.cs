namespace MASICResultsMerger
{
    internal class ScanStatsData
    {
        public int ScanNumber { get; }
        public string ElutionTime { get; set; }
        public string ScanType { get; set; }
        public string TotalIonIntensity { get; set; }
        public string BasePeakIntensity { get; set; }
        public string BasePeakMZ { get; set; }

        /// <summary>
        /// Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        /// </summary>
        public string CollisionMode { get; set; }

        /// <summary>
        /// Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        /// </summary>
        public string ReporterIonData { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="scanNumber"></param>
        public ScanStatsData(int scanNumber)
        {
            ScanNumber = scanNumber;
        }

        /// <summary>
        /// Compare the scan numbers between two scan stats items
        /// </summary>
        /// <param name="other">Scan stats instance to compare</param>
        /// <returns>Comparison result</returns>
        public int CompareTo(ScanStatsData other)
        {
            if (ScanNumber < other.ScanNumber)
            {
                return -1;
            }

            if (ScanNumber > other.ScanNumber)
            {
                return 1;
            }

            return 0;
        }
    }
}
