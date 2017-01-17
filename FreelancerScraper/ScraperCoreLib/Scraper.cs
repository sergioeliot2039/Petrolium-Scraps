using ScraperModel;

namespace ScraperCoreLib
{
    public class Scraper
    {
        public static void Run(IGrabData grabber, IScrape scraper, string path)
        {
            var grabbedData = grabber.Grab(path);
            using (var db = new ScraperDbContext())
            {
                scraper.ScrapeData(db, grabbedData);
            }
        }

    }
}
