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
    public class GovUkOilProductionAndTradeScraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/548011/Oil_Production___Trade_since_1890.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Oli Production And Trade 
            var oilProductionAndTradeSheet = xls.GetSheet("Production, Imports & Exports");
            var oilProductionAndTradeData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int firstRowWithData = 11;
            for (int counter = firstRowWithData; counter < 149; counter++)
            {
                #region Scrap logic - Imports of Oil and Oil Products
                if (oilProductionAndTradeSheet.GetRow(counter) != null)
                {
                    for (int cellIdx = oilProductionAndTradeSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx < oilProductionAndTradeSheet.GetRow(counter).LastCellNum; cellIdx++)
                    {
                        GovUkSupplyUseCrudeGasFeedstocks oilProductionAndTradeDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                        oilProductionAndTradeDatum.Source = URL;

                        int year = int.Parse(oilProductionAndTradeSheet.GetRow(counter).GetCell(oilProductionAndTradeSheet.GetRow(counter).FirstCellNum).NumericCellValue.ToString());
                        if (year>=1890)//Skip empty rows
                        { 
                            oilProductionAndTradeDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                            oilProductionAndTradeDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                            var currentCell = oilProductionAndTradeSheet.GetRow(counter).GetCell(cellIdx);
                            if (currentCell != null)
                            {
                                oilProductionAndTradeDatum.Quantity = currentCell.NumericCellValue;

                                switch (cellIdx)
                                {

                                    case 1: oilProductionAndTradeDatum.Name = "Crude Oil Imports"; break;
                                    case 2: oilProductionAndTradeDatum.Name = "Crude Oil Ind. Production Total"; break;
                                    case 3: oilProductionAndTradeDatum.Name = "Crude Oil Ind. Production Landward"; break;
                                    case 4: oilProductionAndTradeDatum.Name = "Crude Oil Ind. Production Feed stocks"; break;
                                    case 5: oilProductionAndTradeDatum.Name = "Crude Oil Exports"; break;
                                    case 6: oilProductionAndTradeDatum.Name = "Crude Oil Refinery throughput"; break;
                                    case 8: oilProductionAndTradeDatum.Name = "Oil products Refinery output"; break;
                                    case 9: oilProductionAndTradeDatum.Name = "Oil products Exports"; break;
                                    case 10: oilProductionAndTradeDatum.Name = "Oil products Imports"; break;
                                    case 11: oilProductionAndTradeDatum.Name = "Oil products Inland deliveries"; break;
                                    case 14: oilProductionAndTradeDatum.Name = "Net Exports Crude Oil"; break;
                                    case 15: oilProductionAndTradeDatum.Name = "Net Exports Oil Products"; break;
                                    case 16: oilProductionAndTradeDatum.Name = "Net Exports Total"; break;
                                    case 18: oilProductionAndTradeDatum.Name = "Crude Oil Ratio of imports to ref. throughput"; break;
                                    case 19: oilProductionAndTradeDatum.Name = "Crude Oil Ratio of indigenious production to ref. throughput"; break;
                                    case 20: oilProductionAndTradeDatum.Name = "Crude Oil Ratio of exports to indigenious production"; break;
                                    case 22: oilProductionAndTradeDatum.Name = "Oil products Imports: Share of inland deliveries"; break;
                                    default: oilProductionAndTradeDatum.Name = string.Empty; break;
                                }

                                //if (year == 1891 || year == 1950 || year == 2000)
                                //{
                                    if (!string.IsNullOrEmpty(oilProductionAndTradeDatum.Name))
                                        oilProductionAndTradeData.Add(oilProductionAndTradeDatum);
                                        //oilProductionAndTradeData.Add(oilProductionAndTradeDatum);
                                //}
                            }
                        }
                    }
                }
                #endregion
            }
            dbContext.SaveData(oilProductionAndTradeData, lastModified);
            Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in oilProductionAndTradeData) { Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.StartDate.Year, "-"); }

            // Oli Production And Trade 
            var petroleumSheet = xls.GetSheet("Inland Deliveries of Products");
            var petroleumData1 = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift1 = 1;
            int firstRowWithData1 = 8;
            for (int counter = firstRowWithData1; counter < 169; counter++)
            {
                #region Scrap logic - Imports of Oil and Oil Products
                if (petroleumSheet.GetRow(counter) != null)
                {
                    for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift1; cellIdx < petroleumSheet.GetRow(counter).LastCellNum; cellIdx++)
                    {
                        GovUkSupplyUseCrudeGasFeedstocks petroleumDatum1 = new GovUkSupplyUseCrudeGasFeedstocks();
                        petroleumDatum1.Source = URL;

                        int year = int.Parse(petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).NumericCellValue.ToString());
                        if (year >= 1870)//Skip empty rows
                        {
                            petroleumDatum1.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                            petroleumDatum1.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                            var currentCell = petroleumSheet.GetRow(counter).GetCell(cellIdx);
                            if (currentCell != null)
                            {
                                if (currentCell.CellType == CellType.Numeric)
                                    petroleumDatum1.Quantity = currentCell.NumericCellValue;
                                else
                                {
                                    if (currentCell.CellType == CellType.String)
                                    {
                                        string value = currentCell.StringCellValue.Replace("'","").Trim();
                                        try
                                        {
                                            petroleumDatum1.Quantity = int.Parse(value);
                                        }
                                        catch
                                        {
                                            petroleumDatum1.Quantity = 0.0;
                                        }
                                    }
                                    else
                                    {
                                        petroleumDatum1.Quantity = 0.0;
                                    }
                                }

                                switch (cellIdx)
                                {

                                    case 2: petroleumDatum1.Name = "Gases Butane & Propane"; break;
                                    case 3: petroleumDatum1.Name = "Gases Other Petroleum Gases"; break;
                                    case 5: petroleumDatum1.Name = "Feedstocks for Petchem Gases"; break;
                                    case 6: petroleumDatum1.Name = "Feedstocks for Petchem  Naphtha (LDF)"; break;
                                    case 7: petroleumDatum1.Name = "Feedstocks for Petchem  Other Products"; break;
                                    case 8: petroleumDatum1.Name = "Feedstocks for Petchem Total"; break;                                    
                                    case 10: petroleumDatum1.Name = "Naptha (LDF) for Gasworks"; break;
                                    case 11: petroleumDatum1.Name = "Aviation Spirit"; break;
                                    case 12: petroleumDatum1.Name = "Wide Cut Gasoline"; break;
                                    case 13: petroleumDatum1.Name = "Column 13"; break;
                                    case 14: petroleumDatum1.Name = "Motor Spirit"; break;
                                    case 15: petroleumDatum1.Name = "Industrial Spirit"; break;
                                    case 16: petroleumDatum1.Name = "White Spirit"; break;
                                    case 17: petroleumDatum1.Name = "Column 17"; break;
                                    case 18: petroleumDatum1.Name = "Kerosene Aviation Turbine Fuel"; break;
                                    case 19: petroleumDatum1.Name = "Kerosene Burning Oil"; break;
                                    case 20: petroleumDatum1.Name = "Kerosene Vaporising"; break;
                                    case 21: petroleumDatum1.Name = "Column 21"; break;
                                    case 22: petroleumDatum1.Name = "DERV FUEL"; break;
                                    case 23: petroleumDatum1.Name = "Gas Oil"; break;
                                    case 24: petroleumDatum1.Name = "Marine Diesel Oil"; break;
                                    case 25: petroleumDatum1.Name = "Fuel Oil"; break;
                                    case 26: petroleumDatum1.Name = "Lubricating Oil"; break;
                                    case 27: petroleumDatum1.Name = "Bitumen"; break;
                                    case 28: petroleumDatum1.Name = "Paraffin Wax"; break;
                                    case 29: petroleumDatum1.Name = "Petroleum"; break;
                                    case 30: petroleumDatum1.Name = "Misc. Products"; break;
                                    case 31: petroleumDatum1.Name = "TOTAL PRODUCTS"; break;
                                    case 32: petroleumDatum1.Name = "Refinery Fuel"; break;
                                    case 33: petroleumDatum1.Name = "TOTAL (inc Refinery Fuel)"; break;
                                    default: petroleumDatum1.Name = string.Empty; break;
                                }

                                //if (year == 1945 || year == 1957 || year == 1934 || year == 1938 || year == 1920)
                                //{
                                    if (!string.IsNullOrEmpty(petroleumDatum1.Name))
                                        petroleumData1.Add(petroleumDatum1);
                                    //petroleumData1.Add(petroleumDatum1);
                                //}
                            }
                        }
                    }
                }
                #endregion
            }
            dbContext.SaveData(petroleumData1, lastModified);
            Console.WriteLine("Name \t\t Quantity \t Year");

            foreach (GovUkSupplyUseCrudeGasFeedstocks s in petroleumData1) { Console.WriteLine("{0} \t {1} \t {2}", s.Name, s.Quantity, s.StartDate.Year); 
            
            }
        }
    }
}
