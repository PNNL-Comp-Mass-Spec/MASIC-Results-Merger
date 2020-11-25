using System.Collections.Generic;

namespace MASICResultsMerger
{
    class ProcessedFileInfo
    {
        public const string COLLISION_MODE_NOT_DEFINED = "Collision_Mode_Not_Defined";

        public string BaseName { get; set; }

        /// <summary>
        /// The Key is the collision mode and the value is the output file path
        /// </summary>
        public Dictionary<string, string> OutputFiles { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="baseDatasetName"></param>
        public ProcessedFileInfo(string baseDatasetName)
        {
            BaseName = baseDatasetName;
            OutputFiles = new Dictionary<string, string>();
        }

        /// <summary>
        /// Add an output file
        /// </summary>
        /// <param name="collisionMode"></param>
        /// <param name="outputFilePath"></param>
        public void AddOutputFile(string collisionMode, string outputFilePath)
        {
            if (string.IsNullOrEmpty(collisionMode))
            {
                OutputFiles.Add(COLLISION_MODE_NOT_DEFINED, outputFilePath);
            }
            else
            {
                OutputFiles.Add(collisionMode, outputFilePath);
            }
        }
    }
}
