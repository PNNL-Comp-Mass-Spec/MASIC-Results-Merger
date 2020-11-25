
namespace MASICResultsMerger
{
    class ScanStatsData
    {
        public int ScanNumber { get; }
        public string ElutionTime { get; set; }
        public string ScanType { get; set; }
        public string TotalIonIntensity { get; set; }
        public string BasePeakIntensity { get; set; }
        public string BasePeakMZ { get; set; }
        public string CollisionMode { get; set; }

        /// <summary>
        /// Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        /// </summary>
        public string ReporterIonData { get; set; }

        /// <summary>
        /// Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        /// </summary>
        public string ScanStatsFileName { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="scanNumber"></param>
        public ScanStatsData(int scanNumber)
        {
            ScanNumber = scanNumber;
        }

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
