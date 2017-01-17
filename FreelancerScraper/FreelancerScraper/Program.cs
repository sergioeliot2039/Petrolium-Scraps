using ScraperCoreLib;
using ScraperFreelancerLib;

namespace FreelancerScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            // Run getting data remotely
            //Scraper.Run(new WebGrabData(), new GovUkSupplyUseCrudeGasFeedstocksScraper(), GovUkSupplyUseCrudeGasFeedstocksScraper.URL);

            // Run getting data locally (means you can download the file then develop against it without having to keep hitting the web site or if you are offline)
            Scraper.Run(new LocalFileGrabData(), new GovUSAIntlPetroleumNOthersSupplyDispositionNPrices(), @"C:\Users\sergi_000\OneDrive\AT\Petrolium Scraps\Stage 1.5 USA files\1.5.2 Annual Projections\aeotab_21.xlsx");
            
            //
            /* ""*/
            /*
             Scraper.Run(new LocalFileGrabData(), new GovUkAdditionalInfoOnInlandDeliveries4NonenergyUses(), @);
             */
            
            
            
        }
    }
}
