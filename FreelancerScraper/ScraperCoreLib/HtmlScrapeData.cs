using System;
using System.Text;
using ScraperModel;
using HtmlAgilityPack;

namespace ScraperCoreLib
{
    abstract public class HtmlScrapeData : IScrape
    {
        abstract protected void ScrapeHtml(ScraperDbContext dbContext, DateTime lastModified, HtmlDocument htmlDoc);

        public void ScrapeData(ScraperDbContext dbContext, GrabbedData data)
        {
            var dataString = Encoding.UTF8.GetString(data.Data);
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(dataString);
            ScrapeHtml(dbContext, data.LastModifiedTimestamp, htmlDoc);
        }

    }
}
