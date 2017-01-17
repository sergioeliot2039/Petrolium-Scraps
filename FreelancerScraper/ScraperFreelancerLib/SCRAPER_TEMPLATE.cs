using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScraperCoreLib;
using ScraperModel.Models;
using ScraperModel;

namespace ScraperFreelancerLib
{
    public class a : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/541285/DUKES_F.4.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {

        }

    }
}