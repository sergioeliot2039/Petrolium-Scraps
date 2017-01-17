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
    public class GovUkOilCommodityBalance3dot2TO3dot4Scraper : ExcelXlsScrapeData
    {
        public const string URL = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/540914/DUKES_3.2-3.4.xls";

        protected override void ScrapeXls(ScraperDbContext dbContext, DateTime lastModified, IWorkbook xls)
        {
           string[] years = {
                                "1998",  "1999",  "2000",  "2001",  "2002",  "2003",  "2004",  "2005",  "2006", 
                                "2007",  "2008","2009",  "2010",  "2011",  "2012",  "3.4",  "3.3",  "3.2"
                             };
            var petroleumData = new List<GovUkSupplyUseCrudeGasFeedstocks>();
            int cellShift = 1;
            int lastCellWithData = 8;
            int firstRowWithData = 7;
            int rowCellHeaders = 3;
            int lastRow = 65;

            foreach(string year in years)
            {
                string syear = year;
                switch (syear)
                {
                    case "3.4":
                        syear = "2013";
                        break;
                    case "3.3":
                        syear = "2014";
                        break;
                    case "3.2":
                        syear = "2015";
                        break;
                }
                var petroleumSheet = xls.GetSheet(year);

                for (int counter = firstRowWithData; counter <= lastRow; counter++)
                {

                
                    if (petroleumSheet.GetRow(counter) != null)
                    {
                        for (int cellIdx = petroleumSheet.GetRow(counter).FirstCellNum + cellShift; cellIdx <= lastCellWithData; cellIdx++)
                        {



                            GovUkSupplyUseCrudeGasFeedstocks petroleumDatum = new GovUkSupplyUseCrudeGasFeedstocks();
                            petroleumDatum.Source = URL;

                            petroleumDatum.StartDate = new DateTime(int.Parse(syear.ToString()), 1, 1);
                            petroleumDatum.EndDate = new DateTime(int.Parse(syear.ToString()), 12, 31);
                            var currentCell = petroleumSheet.GetRow(counter).GetCell(cellIdx);

                            var superextraRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders + 2).GetCell(cellIdx);
                            var extraRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders + 1).GetCell(cellIdx);
                            var regularRowHeaderCell = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx);
                                
                            if (currentCell != null)
                            {

                                petroleumDatum.Name
                                    = (extraRowHeaderCell != null
                                    && extraRowHeaderCell.CellType == CellType.String)
                                    ? regularRowHeaderCell.StringCellValue + " " + extraRowHeaderCell.StringCellValue 
                                    : regularRowHeaderCell.StringCellValue;

                                if(superextraRowHeaderCell!=null&&superextraRowHeaderCell.CellType == CellType.String)
                                {
                                    petroleumDatum.Name = petroleumDatum.Name 
                                        + " " + superextraRowHeaderCell.StringCellValue 
                                        + " " + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue;
                                }
                                else
                                {
                                    petroleumDatum.Name = petroleumDatum.Name 
                                        + " " + petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue;
                                }


                                if(petroleumDatum.Name.Contains("Total") && petroleumDatum.Name.Contains("Products"))
                                    break;


                                petroleumDatum.Quantity = currentCell.CellType == CellType.Numeric ?  currentCell.NumericCellValue : 0;

                                int yeara = int.Parse(syear);
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