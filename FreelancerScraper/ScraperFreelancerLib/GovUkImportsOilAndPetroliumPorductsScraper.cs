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
    public class GovUkImportsOilAndPetroliumProductsScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/548012/Oil_Imports_since_1920.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Imports of Oil and Oil Products
            var importsOfOilProductsSheet = xls.GetSheet("Imports of Oil and Oil Products");
            var importsOfOilProductsData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 2;
            for (int counter = 9; counter < 105; counter++)
            {
                #region Scrap logic - Imports of Oil and Oil Products
                if (importsOfOilProductsSheet.GetRow(counter) != null)
                {
                    for (int cellIdx = importsOfOilProductsSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx < importsOfOilProductsSheet.GetRow(counter).LastCellNum; cellIdx++)
                    {
                        GovUkSupplyUseCrudeGasFeedstocks importsOfOilProductsDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                        importsOfOilProductsDatum.Source = URL;

                        int year = int.Parse(importsOfOilProductsSheet.GetRow(counter).GetCell(importsOfOilProductsSheet.GetRow(counter).FirstCellNum).NumericCellValue.ToString());
                        importsOfOilProductsDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        importsOfOilProductsDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        var currentCell = importsOfOilProductsSheet.GetRow(counter).GetCell(cellIdx);
                        if (currentCell!=null)
                        {
                            importsOfOilProductsDatum.Quantity = currentCell.NumericCellValue;

                            switch (cellIdx)
                            {     

                                case 2: importsOfOilProductsDatum.Name  = "Crude Oil - Crude Oil and process oils"; break;
                                case 5: importsOfOilProductsDatum.Name  = "Motor Spirit"; break;
                                case 6: importsOfOilProductsDatum.Name  = "Aviation Spirit"; break;
                                case 7: importsOfOilProductsDatum.Name  = "Other Spirits";
                                    if (importsOfOilProductsDatum.Quantity==0.0)
                                        importsOfOilProductsDatum.Quantity = importsOfOilProductsSheet.GetRow(counter).GetCell(cellIdx+1).NumericCellValue;
                                            break;
                                
                                case 10: importsOfOilProductsDatum.Name = "Naphtha"; break;
                                case 12: importsOfOilProductsDatum.Name = "Aviation Turbine Fuel"; break;
                                case 13: importsOfOilProductsDatum.Name = "Burning Oil"; break;
                                case 15: importsOfOilProductsDatum.Name = "Gas Oil"; break;
                                case 16: importsOfOilProductsDatum.Name = "Diesel Oil";break;
                                case 17: importsOfOilProductsDatum.Name = "Fuel Oil"; break;
                                case 18: importsOfOilProductsDatum.Name = "Lubricating Oil"; break;
                                case 19: importsOfOilProductsDatum.Name = "Misc. Products"; break;
                                case 20: importsOfOilProductsDatum.Name = "Total Products"; break;
                                default: importsOfOilProductsDatum.Name = string.Empty; break;
                              }

                            //if (year == 1944 || year == 1969 || year == 1994)
                            //{
                                if (!string.IsNullOrEmpty(importsOfOilProductsDatum.Name))
                                    importsOfOilProductsData.Add(importsOfOilProductsDatum);
                            //}
                        }
                    }
                }
                #endregion
            }
            dbContext.SaveData(importsOfOilProductsData, lastModified);
            Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in importsOfOilProductsData){Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.StartDate.Year, "-");}

            // Imports of Crude Oil by Country
            var importsOfOilByCountrySheet = xls.GetSheet("Imports of Crude Oil by Country");
            var importsOfOilByCountryData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            cellShift = 1;
            int rowCellHeaders = 4;
            for (int counter = 5; counter < 72; counter++)
            {
                for (int cellIdx = importsOfOilByCountrySheet.GetRow(counter).FirstCellNum + cellShift; cellIdx < importsOfOilByCountrySheet.GetRow(counter).LastCellNum; cellIdx++)
                {
                        GovUkSupplyUseCrudeGasFeedstocks importsOfOilByCountryDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                        importsOfOilByCountryDatum.Source = URL;

                        int year = int.Parse(importsOfOilByCountrySheet.GetRow(counter).GetCell(importsOfOilByCountrySheet.GetRow(counter).FirstCellNum).NumericCellValue.ToString());
                        importsOfOilByCountryDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        importsOfOilByCountryDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        var currentCell = importsOfOilByCountrySheet.GetRow(counter).GetCell(cellIdx);
                        if (currentCell != null)
                        {
                            importsOfOilByCountryDatum.Quantity = currentCell.NumericCellValue;
                        }
                        importsOfOilByCountryDatum.Name = importsOfOilByCountrySheet.GetRow(rowCellHeaders).GetCell(cellIdx).StringCellValue;
                    //if(year==1938)
                    //{ 
                        if (!string.IsNullOrEmpty(importsOfOilByCountryDatum.Name))
                            importsOfOilByCountryData.Add(importsOfOilByCountryDatum);
                    //}

                }
            }
            dbContext.SaveData(importsOfOilByCountryData, lastModified);
            Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in importsOfOilByCountryData) { Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.StartDate.Year, "-"); }
            Console.WriteLine("Done");
            Console.ReadKey();

        }

    }
}
