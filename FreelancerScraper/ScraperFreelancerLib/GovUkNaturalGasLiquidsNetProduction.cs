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
    public class GovUkNaturalGasLiquidsNetProduction : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/541284/DUKES_F.3.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet F.3
            var petroleumSheet = xls.GetSheet("F.3");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 2;
            int rowCellHeaders = 2;
            int lastCellWithData = 18;
            int lastRow = 28;
            int firstRowWithData = 4;
            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                for (int cellIdx = petroleumSheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (rowIdx == 9 || rowIdx == 15 || rowIdx == 18)
                        continue;
                    GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                    petroleumDatum.Source = URL;
                    var currentCell = petroleumSheet.GetRow(rowIdx).GetCell(cellIdx);
                    if (currentCell != null)
                    {


                        double year = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue;

                        petroleumDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        petroleumDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        if (currentCell.CellType == CellType.Numeric || currentCell.CellType == CellType.Formula)
                        {
                            petroleumDatum.Quantity = currentCell.NumericCellValue;
                        }
                        else
                        {
                            petroleumDatum.Quantity = 0.0;
                        }


                        string prefixName = "";
                        if (rowIdx >= 4 && rowIdx <= 7)
                        {
                            prefixName = "Offshore oil pipeline terminals (1): ";
                        }
                        else if (rowIdx >= 10 && rowIdx <= 13)
                        {
                            prefixName = "Offshore associated gas terminals (2): ";
                        }
                        else if (rowIdx == 16)
                        {
                            prefixName = "Offshore dry gas terminals (3): ";
                        }
                        else if (rowIdx >= 19 && rowIdx <= 22)
                        {
                            prefixName = "Onshore production (4): ";
                        }

                            
                        petroleumDatum.Name = prefixName + petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum).StringCellValue;

                        /*try
                        {
                            Console.WriteLine(petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).StringCellValue);
                        }
                        catch { Console.WriteLine(petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue); }*/
                        /*if (petroleumDatum.Name.Contains("final") && petroleumDatum.Name.Contains("users"))
                            break;*/



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