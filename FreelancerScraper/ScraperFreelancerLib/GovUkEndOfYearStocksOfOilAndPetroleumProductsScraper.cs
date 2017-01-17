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
    public class GovUkEndOfYearStocksOfOilAndPetroleumProductsScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/540918/DUKES_3.7.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet 3.7
            var petroleumSheet = xls.GetSheet("3.7");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int rowCellHeaders = 4;
            int lastCellWithData = 19;
            int lastRow = 32;
            int firstRowWithData = 6;

            for (int counter = firstRowWithData; counter <= lastRow; counter++)
            {

                for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (counter == 5 || counter == 11 || counter == 12)
                        break;

                    GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                    petroleumDatum.Source = URL;
                    var currentCell = petroleumSheet.GetRow(counter).GetCell(cellIdx);
                    if (currentCell != null)
                    {

                        double year = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue;
                        petroleumDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        petroleumDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        if (currentCell.CellType == CellType.Numeric)
                        {
                            petroleumDatum.Quantity = currentCell.NumericCellValue;
                        }
                        else
                        {
                            petroleumDatum.Quantity = 0.0;
                        }

                        string prefixName = "";
                        if (counter >= 6 && counter <= 9)
                        {
                            prefixName = "Crude and process oils ";
                        }
                        else if (counter >= 13 && counter <= 30)
                        {
                            prefixName = "Petroleum products ";
                        }
                        else if (counter == 32 )
                        {
                            prefixName = "Total all products ";
                        }
                        else if (counter == 24)
                        {
                            prefixName = "Gas/Diesel oil (7) ";
                        }
                        
                        petroleumDatum.Name = prefixName
                                    + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue;

                        if (year == 1999)
                        {
                            if (!string.IsNullOrEmpty(petroleumDatum.Name))
                                petroleumData.Add(petroleumDatum);
                        }
                    }
                }
            }
            dbContext.SaveData(petroleumData, lastModified);
            Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in petroleumData) { Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.StartDate.Year, "-"); }
            Console.WriteLine("Done");
            Console.ReadKey();
        }

    }
}