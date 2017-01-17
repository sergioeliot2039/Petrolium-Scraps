using System;
using ScraperModel;
using System.IO;
using ClosedXML.Excel;

namespace ScraperCoreLib
{
    abstract public class ExcelXlsxScrapeData : IScrape
    {
        abstract protected void ScrapeXlsx(ScraperDbContext dbContext, DateTime lastModified, XLWorkbook xlsx);

        public void ScrapeData(ScraperDbContext dbContext, GrabbedData data)
        {
            using (var stream = new MemoryStream(data.Data))
            {
                using (var xlsx = new XLWorkbook(stream, XLEventTracking.Disabled))
                {
                    ScrapeXlsx(dbContext, data.LastModifiedTimestamp, xlsx);
                }
            }
        }
    }
}
