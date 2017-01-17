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
    public class GovUkPetroleumSupplayAndDisposal : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/540916/DUKES_3.5.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Sheet 3.5
            var petroleumSheet = xls.GetSheet("3.5");
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int rowCellHeaders = 3;
            int lastCellWithData = 19;
            int lastRow = 33;
            int firstRowWithData = 5;
            for (int counter = 6; counter <= lastRow; counter++)
            {

                if (counter == 29 || counter == 16)
                    continue;

                for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
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
                    if(counter>5 && counter <= 14)
                    {
                        prefixName = "Primary oils (Crude oil, NGLs and feedstocks) "; 
                    }
                    else if (counter >= 17 && counter <= 28)
                    {
                        prefixName = "Petroleum products "; 
                    }
                    else if (counter == 31 )
                    {
                        prefixName = "Energy use ";
                    }
                    else if (counter == 32)
                    {
                        prefixName = "Energy use Of which, ";
                    }

                    petroleumDatum.Name = prefixName 
                                + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue;   

                    if (year == 1997)
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