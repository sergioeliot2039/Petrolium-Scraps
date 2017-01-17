using System.Net;
using System.IO;

namespace ScraperCoreLib
{
    public class WebGrabData : IGrabData
    {
        public GrabbedData Grab(string path)
        {
            var grabbedData = new GrabbedData();
            HttpWebRequest req = (HttpWebRequest) HttpWebRequest.Create(path );
            req.CookieContainer = new CookieContainer();
            req.Headers.Add("pragma", "no-cache");
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0";
            req.Proxy = WebRequest.DefaultWebProxy;
            req.UseDefaultCredentials = true;
            using (var httpResponse = (HttpWebResponse)req.GetResponse())
            {
                grabbedData.LastModifiedTimestamp = httpResponse.LastModified;
                var length = httpResponse.ContentLength;
                using (var stream = httpResponse.GetResponseStream())
                {
                    using (var reader = new StreamReader(stream))
                    {
                        if (length < 0)
                        {
                            using (var memory = new MemoryStream())
                            {
                                stream.CopyTo(memory);
                                grabbedData.Data = memory.ToArray();
                            }
                        }
                        else
                        {
                            var bytes = new byte[length];
                            int position = 0;
                            while (position < length)
                            {
                                int bytesRead = stream.Read(bytes, position, (int)length - position);
                                if (bytesRead == 0)
                                    throw new IOException(@"Premature end of data at position[{position}]. Stream of length[{length}]");

                                position += bytesRead;
                            }
                            grabbedData.Data = bytes;
                        }
                    }
                }
            }
            return grabbedData;
        }
    }
}
