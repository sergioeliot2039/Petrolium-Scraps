using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using ScraperCoreLib;
using ScraperModel.Models;
using ScraperModel;

namespace ScraperFreelancerLib
{
    public class GovUkSupplyUseCrudeGasFeedstocksScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/556229/ET_3.1.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Annual data
            var annualSheet = xls.GetSheet("Annual");
            var annualYearRow = annualSheet.GetRow(4);
            var annualData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            for (int counter=5; counter<annualSheet.LastRowNum; counter++)
            {
                // Create GovUkSupplyUseCrudeGasFeedstocks objects for annual data
            }
            dbContext.SaveData(annualData, lastModified);


            // Quarterly data
            var quarterSheet = xls.GetSheet("Quarter");
            var quarterYearRow = quarterSheet.GetRow(4);
            var quarterRow = quarterSheet.GetRow(5);
            var quarterData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            for (int counter = 6; counter < quarterSheet.LastRowNum; counter++)
            {
                // Create GovUkSupplyUseCrudeGasFeedstocks objects for quarterly data
            }
            dbContext.SaveData(quarterData, lastModified);
        }

    }

}
