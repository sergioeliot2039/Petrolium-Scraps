using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScraperCoreLib;
using ScraperModel.Models;
using ScraperModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ScraperFreelancerLib
{
    public class GovUkCrudeOilAndNaturalGasProduction : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/541283/DUKES_F.1.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet F.1
            var petroleumSheet = xls.GetSheet("F.1");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 2;
            int rowCellHeaders = 2;
            int lastCellWithData = 21;
            int lastRow = 17;
            int firstRowWithData = 3;
            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                for (int cellIdx = petroleumSheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (rowIdx == 14 || rowIdx == 16)
                        continue;
                    GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                    petroleumDatum.Source = URL;
                    var currentCell = petroleumSheet.GetRow(rowIdx).GetCell(cellIdx);
                    if (currentCell != null)
                    {


                        double year = 0;
                        if (petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).CellType == CellType.String)
                        {
                            Regex regex = new Regex("[0-9]+");
                            MatchCollection m = regex.Matches(petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            {
                                year = double.Parse(ma.Value);
                                break;
                            }
                        }
                        else
                        {
                            year = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue;
                        }
                        
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

                        string totalPrefix = "";
                        if (cellIdx == 2 || cellIdx == 21)
                            totalPrefix = "Total to end ";

                        string crudeOilPrefix = "";
                        if (rowIdx >= 3 && rowIdx <= 12)
                        {
                            crudeOilPrefix = "CRUDE OIL ";
                        }

                        string prefixName = "";
                        if (rowIdx == 3 )
                        {
                            prefixName = "Offshore production: ";
                        }
                        if (rowIdx >= 4 && rowIdx <= 9)
                        {
                            prefixName = "Terminal Receipts: ";
                        }
                        else if (rowIdx >= 11 && rowIdx <= 12)
                        {
                            prefixName = "Total Terminal Receipts: ";
                        }

                        if (rowIdx == 13 || rowIdx == 15 || rowIdx == 17)
                        {
                            petroleumDatum.Name = totalPrefix + petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum).StringCellValue;
                        }
                        else
                        {
                            if (rowIdx >= 11 && rowIdx <= 12 || rowIdx==10)
                                petroleumDatum.Name = crudeOilPrefix + totalPrefix + prefixName+ petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum).StringCellValue+ petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;                       
                            else
                                petroleumDatum.Name = crudeOilPrefix + totalPrefix + prefixName+ petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;                       
                        }
                        /*try
                        {
                            Console.WriteLine(petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).StringCellValue);
                        }
                        catch { Console.WriteLine(petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue); }*/
                        /*if (petroleumDatum.Name.Contains("final") && petroleumDatum.Name.Contains("users"))
                            break;*/



                        if (year == 2015 && !petroleumDatum.Name.Contains("end"))
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