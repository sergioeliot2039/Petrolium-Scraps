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
    public class GovUkOilCommodityBalanceScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/540912/DUKES_3.1_web.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {

            string[] years = {
                                "1998",  "1999",  "2000",  "2001",  "2002",  "2003",  "2004",  "2005",  "2006", 
                                "2007",  "2008","2009",  "2010",  "2011",  "2012",  "2013",  "2014",  "2015"
                             };
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int lastCellWithData = 8;
            int firstRowWithData = 5;
            int rowCellHeaders = 3;
            int lastRow = 37;

            foreach(string year in years)
            {
                var petroleumSheet = xls.GetSheet(year);

                for (int counter = firstRowWithData; counter <= lastRow; counter++)
                {

                
                    if (petroleumSheet.GetRow(counter) != null)
                    {
                        for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                        {


                            GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                            petroleumDatum.Source = URL;

                            petroleumDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                            petroleumDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                            var currentCell = petroleumSheet.GetRow(counter).GetCell(cellIdx);

                            var extraRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders + 1).GetCell(cellIdx);
                            var regularRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx);
                                
                            if (currentCell != null)
                            {
                                petroleumDatum.Name
                                 = (extraRowHeaderCell != null
                                    && extraRowHeaderCell.CellType == CellType.String)
                                    ? regularRowHeaderCell.StringCellValue + " " + extraRowHeaderCell.StringCellValue + " " + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue
                                    : regularRowHeaderCell.StringCellValue + " " + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue;                                

                                petroleumDatum.Quantity = currentCell.CellType == CellType.Numeric ?  currentCell.NumericCellValue : 0;

                                int yeara = int.Parse(year);
                                if (yeara == 2005)
                                {
                                    if (!string.IsNullOrEmpty(petroleumDatum.Name))
                                    {
                                        petroleumData.Add(petroleumDatum);

                                    }
                                    //petroleumData.Add(petroleumDatum);
                                }
                            }
                        }
                    }
                
                }//iterates rows
            }//iterates sheetnames (their names are years)
            dbContext.SaveData(petroleumData, lastModified);

            Console.WriteLine("Name \t\t Quantity \t Year");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in petroleumData) { Console.WriteLine("{0} \t {1} \t {2}", 
                s.Name, s.Quantity, s.StartDate.Year); }
            Console.ReadKey();
        }

    }
}