namespace MASICResultsMerger
{
    internal class MASICFileInfo
    {
        public string ScanStatsFileName { get; set; }
        public string SICStatsFileName { get; set; }
        public string ReporterIonsFileName { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public MASICFileInfo()
        {
            ScanStatsFileName = string.Empty;
            SICStatsFileName = string.Empty;
            ReporterIonsFileName = string.Empty;
        }
    }
}
