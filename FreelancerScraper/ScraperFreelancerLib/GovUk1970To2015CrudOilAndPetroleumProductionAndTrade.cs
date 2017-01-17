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
    public class GovUk1970To2015CrudOilAndPetroleumProductionAndTrade : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/541280/DUKES_3.1.1.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
            // Crud Oli and Oil products Trade  from 1970 To 2015
            var oilProductionAndTradeSheet = xls.GetSheet("3.1.1");
            var oilProductionAndTradeData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int firstRowWithData = 11;
            int lastRow = 56;
            int lastCellWithData = 22;

            for (int counter = firstRowWithData; counter <= lastRow; counter++)
            {
                #region Scrap logic - Imports of Oil and Oil Products
                if (oilProductionAndTradeSheet.GetRow(counter) != null)
                {
                    for (int cellIdx = oilProductionAndTradeSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                    {
                        GovUkSupplyUseCrudeGasFeedstocks oilProductionAndTradeDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                        oilProductionAndTradeDatum.Source = URL;

                        int year = int.Parse(oilProductionAndTradeSheet.GetRow(counter).GetCell(oilProductionAndTradeSheet.GetRow(counter).FirstCellNum).NumericCellValue.ToString());

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

                            if (year == 2015)
                            {
                                if (!string.IsNullOrEmpty(oilProductionAndTradeDatum.Name))
                                    oilProductionAndTradeData.Add(oilProductionAndTradeDatum);
                                //oilProductionAndTradeData.Add(oilProductionAndTradeDatum);
                            }
                        }                        
                    }
                }
                #endregion
            }
            dbContext.SaveData(oilProductionAndTradeData, lastModified);
            Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
            foreach (GovUkSupplyUseCrudeGasFeedstocks s in oilProductionAndTradeData) { Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.StartDate.Year, "-"); }
            Console.ReadKey();

        }
    }
}
