using System;
using ScraperModel;
using System.IO;
using NPOI.SS.UserModel;

namespace ScraperCoreLib
{
    abstract public class ExcelXlsScrapeData : IScrape
    {
        abstract protected void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls);

        public void ScrapeData(ScraperDbContext dbContext, GrabbedData data)
        {
            using (var stream = new MemoryStream(data.Data))
            {
                var xls = WorkbookFactory.Create(stream, ImportOption.All);
                ScrapeXls(dbContext, data.LastModifiedTimestamp, xls);
            }
        }
    }
}
