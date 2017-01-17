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
    public class GovUk1970To2015InlandDeliveriesOfPetroleumScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/541281/DUKES_3.1.2.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet 3.1.2
            var petroleumSheet = xls.GetSheet("3.1.2");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int rowCellHeaders = 4;
            int lastCellWithData = 22;
            int lastRow = 52;
            int firstRowWithData = 7;
            for (int counter = firstRowWithData; counter <= lastRow; counter++)
            {

                for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (cellIdx == 11 || cellIdx == 16)
                        continue;
                    GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                    petroleumDatum.Source = URL;
                    var currentCell = petroleumSheet.GetRow(counter).GetCell(cellIdx);
                    if (currentCell != null)
                    {

                        double year = petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).NumericCellValue;
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
                        if (cellIdx == 1)
                        {
                            prefixName = "Total ";
                        }
                        if (counter >= 2 && counter <= 9)
                        {
                            prefixName = "Deliveries for energy uses ";
                        }
                        else if (counter >= 12 && counter <= 15)
                        {
                            prefixName = "Energy industry use ";
                        }
                        else if (counter >= 17 && counter <= 21)
                        {
                            prefixName = "Final Users ";
                        }

                        var superextraRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders + 2).GetCell(cellIdx);
                        var extraRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders + 1).GetCell(cellIdx);
                        var regularRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx);
                                
                        petroleumDatum.Name
                            = (extraRowHeaderCell != null
                            && extraRowHeaderCell.CellType == CellType.String)
                            ? regularRowHeaderCell.StringCellValue + " " + extraRowHeaderCell.StringCellValue 
                            : regularRowHeaderCell.StringCellValue;

                        if(superextraRowHeaderCell!=null&&superextraRowHeaderCell.CellType == CellType.String)
                        {
                            petroleumDatum.Name = petroleumDatum.Name 
                                + " " + superextraRowHeaderCell.StringCellValue;
                        }

                        petroleumDatum.Name = prefixName + " " + petroleumDatum.Name;

                        if(petroleumDatum.Name.Contains("final") && petroleumDatum.Name.Contains("users"))
                            break;



                        if (year == 2015)
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