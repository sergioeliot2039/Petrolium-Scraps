using ScraperModel;

namespace ScraperCoreLib
{
    public interface IScrape
    {
        void ScrapeData(ScraperDbContext dbContext, GrabbedData data);


    }
}
