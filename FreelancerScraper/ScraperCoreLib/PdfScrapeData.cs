using iTextSharp.text.pdf;
using ScraperModel;
using System;

namespace ScraperCoreLib
{
    abstract public class PdfScrapeData : IScrape
    {
        abstract protected void ScrapePdf(ScraperDbContext dbContext, DateTime lastModified, PdfReader pdfReader);

        public void ScrapeData(ScraperDbContext dbContext, GrabbedData data)
        {
            using (var pdfReader = new PdfReader(data.Data))
            {
                ScrapePdf(dbContext, data.LastModifiedTimestamp, pdfReader);
            }
        }
    }
}
