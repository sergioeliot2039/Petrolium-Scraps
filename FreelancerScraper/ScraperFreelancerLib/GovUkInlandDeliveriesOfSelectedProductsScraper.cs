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
    public class GovUkInlandDeliveriesOfSelectedProductsScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/540917/DUKES_3.6.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet 3.6
            var petroleumSheet = xls.GetSheet("3.6");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 2;
            int rowCellHeaders = 5;
            int lastCellWithData = 18;
            int lastRow = 29;
            int firstRowWithData = 7;
            for (int counter = firstRowWithData; counter <= lastRow; counter++)
            {

                for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (counter == 13 || counter == 21 || counter == 25)
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
                        if (counter > 6 && counter <= 8)
                        {
                            prefixName = "Motor Spirit ";
                        }
                        else if (counter >= 10 && counter <= 12)
                        {
                            prefixName = "Total Motor Spirit including Bio-ethanol ";
                        }
                        else if (counter >= 14 && counter <= 15)
                        {
                            prefixName = "Diesel Road Fuel ";
                        }
                        else if (counter >= 17 && counter <= 19)
                        {
                            prefixName = "Total Diesel Road Fuel including Bio-diesel ";
                        }
                        else if (counter >= 22 && counter <= 24)
                        {
                            prefixName = "Aviation Fuels ";
                        }
                        else if (counter >= 26 && counter <= 29)
                        {
                            prefixName = "Fuel Oil ";
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