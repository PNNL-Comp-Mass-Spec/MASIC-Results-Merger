namespace MASICResultsMerger
{
    class DatasetInfo
    {
        public string DatasetName { get; set; }
        public int DatasetID { get; set; }

        public DatasetInfo(string datasetName, int datasetId)
        {
            DatasetName = datasetName;
            DatasetID = datasetId;
        }
    }
}
