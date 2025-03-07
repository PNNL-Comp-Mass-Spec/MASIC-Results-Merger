namespace MASICResultsMerger
{
    internal class DatasetInfo
    {
        public string DatasetName { get; set; }
        public int DatasetID { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="datasetName">Dataset name</param>
        /// <param name="datasetId">Dataset ID</param>
        public DatasetInfo(string datasetName, int datasetId)
        {
            DatasetName = datasetName;
            DatasetID = datasetId;
        }
    }
}
