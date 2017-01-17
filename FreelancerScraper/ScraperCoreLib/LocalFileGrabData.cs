using System.IO;

namespace ScraperCoreLib
{
    public class LocalFileGrabData : IGrabData
    {
        public GrabbedData Grab(string path)
        {
            var grabbedData = new GrabbedData();
            grabbedData.Data = File.ReadAllBytes(path);
            var fileInfo = new FileInfo(path);
            grabbedData.LastModifiedTimestamp = fileInfo.LastWriteTimeUtc;
            return grabbedData;
        }
    }
}
