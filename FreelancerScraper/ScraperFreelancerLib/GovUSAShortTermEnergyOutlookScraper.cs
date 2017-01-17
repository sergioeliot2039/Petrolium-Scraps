using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScraperCoreLib;
using ScraperModel.Models;
using ScraperModel;
using ClosedXML.Excel;

namespace ScraperFreelancerLib
{
    public class GovUSAShortTermEnergyOutlookScraper : ExcelXlsxScrapeData
    {
        public const string URL = "https://www.eia.gov/outlooks/steo/xls/STEO_m.xlsx";

        protected override void ScrapeXlsx(ScraperDbContext dbContext, DateTime lastModified, XLWorkbook xlsx)
        {

            #region Sheet 1tab
            var petroleumSheet = xlsx.Worksheet("1tab");
            var petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            int firstRowWithData = 8;
            var cellShift = 3;
            var rowCellHeaders = 4;
            int lastCellWithData = 74;
            int lastRow = 69;
            string category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                var previusRow = petroleumSheet.Row(rowIdx - 1);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString()) && string.IsNullOrEmpty(petroleumRow.Cell(2).Value.ToString()) )
                        break;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    string prefixName = "";
                    if (rowIdx >= 8 && rowIdx <= 14)
                        prefixName = "Energy Supply ";
                    if (rowIdx >= 19 && rowIdx <= 34)
                        prefixName = "Energy Consumption   ";
                    if (rowIdx >= 39 && rowIdx <= 45)
                        prefixName = "Energy Prices ";
                    if (rowIdx >= 50 && rowIdx <= 63)
                        prefixName = "Macroeconomic ";
                    if (rowIdx >= 67 && rowIdx <= 69)
                        prefixName = "Weather ";
                    if (string.IsNullOrEmpty(previusRow.Cell(1).Value.ToString()))
                    {
                        petroleumDatum.Name = prefixName + previusRow.Cell(2).Value.ToString() + " " + petroleumRow.Cell(2).Value.ToString();
                    }
                    else
                    {
                        petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    }
                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();

                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    if (rowIdx < 63) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 1tab

            #region Sheet 2tab
            //Sheet 2tab
             petroleumSheet = xlsx.Worksheet("2tab");
             petroleumData = new List<GovUSAShortTermEnergyOutlook>();
             firstRowWithData = 6;
             cellShift = 3;
             rowCellHeaders = 4;
             lastCellWithData = 74;
             lastRow = 39;

            category = "";
            
            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            { 
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 9)
                        prefixName = "Crude Oil (dollars per barrel) ";
                    if (rowIdx >= 12 && rowIdx <= 14)
                        prefixName = "U.S. Liquid Fuels (cents per gallon) Refiner Prices for Resale";
                    if (rowIdx >= 16 && rowIdx <= 17)
                        prefixName = "U.S. Liquid Fuels (cents per gallon) Refiner Prices to End Users ";
                    if (rowIdx >= 19 && rowIdx <= 22)
                        prefixName = "U.S. Liquid Fuels (cents per gallon) Retail Prices Including Taxes ";
                    if (rowIdx >= 24 && rowIdx <= 25)
                        prefixName = "U.S. Liquid Fuels (cents per gallon) Natural Gas ";
                    if (rowIdx >= 27 && rowIdx <= 29)
                        prefixName = "U.S. Liquid Fuels (cents per gallon) U.S. Retail Prices (dollars per thousand cubic feet) ";
                    if (rowIdx >= 32 && rowIdx <= 35)                        
                        prefixName = "U.S. Electricity Power Generation Fuel Costs (dollars per million Btu) ";
                    if (rowIdx >= 37 && rowIdx <= 39)
                        prefixName = "U.S. Electricity Retail Prices (cents per kilowatthour) ";


                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName+petroleumRow.Cell(2).Value.ToString();
                    petroleumDatum.Quantity = double.Parse(petroleumRow.Cell(cellIdx).Value.ToString());

                    DateTime startDt = new DateTime();
                    DateTime endDt   = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(),out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;
                    /*if (!shown)
                    {
                        Console.WriteLine(petroleumDatum.Name);
                        shown = true;
                    }*/

                }

                /*if(rowIdx==7)
                    break;*/
            }
            /*
            Console.WriteLine("Name \t\t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}",
                    s.Name, s.Quantity, s.StartDate.Year+" "+s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion Sheet 2tab

            #region Sheet 3atab
            petroleumSheet = xlsx.Worksheet("3atab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 47;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 20 || rowIdx == 22 || rowIdx == 38 || rowIdx == 44)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 19)
                        prefixName = "Supply (million barrels per day) (a) ";
                    if (rowIdx >= 24 && rowIdx <= 37)
                        prefixName = "Consumption (million barrels per day) (d) ";
                    if (rowIdx >= 40 && rowIdx <= 43)
                        prefixName = "Total Crude Oil and Other Liquids Inventory Net Withdrawals (million barrels per day) ";
                    if (rowIdx >= 46 && rowIdx <= 47)
                        prefixName = "End-of-period Commercial Crude Oil and Other Liquids Inventories ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    petroleumDatum.Quantity = double.Parse(petroleumRow.Cell(cellIdx).Value.ToString());

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;
                    /*
                    if (!shown)
                    {
                        Console.WriteLine(petroleumDatum.Name);
                        shown = true;
                    }
                    */

                }

                /*if(rowIdx!=7)
                    break;*/

                /*if (rowIdx < 46)
                    continue;*/
            }
            
            /*
            Console.WriteLine("Name \t\t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}",
                    s.Name, s.Quantity, s.StartDate.Year+" "+s.StartDate.Month);
            }            
            Console.ReadKey();
            */
            #endregion  Sheet 3atab

            #region Sheet 3btab
            petroleumSheet = xlsx.Worksheet("3btab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 51;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 10 || rowIdx == 16 || rowIdx == 21 || rowIdx == 28 || rowIdx == 33 || rowIdx == 40
                        || rowIdx == 45 || rowIdx == 47 || rowIdx == 50)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    /*if (rowIdx < 49)
                        continue;*/
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;
                    
                    /*if (!shown)
                    {
                        Console.WriteLine(petroleumDatum.Name);
                        shown = true;
                    }*/
                    

                }

                /*if(rowIdx!=7)
                    break;*/


            }

            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }    
            
            Console.ReadKey();
            */
            #endregion  Sheet 3btab

            #region Sheet 3ctab
            petroleumSheet = xlsx.Worksheet("3ctab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 38;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 21 || rowIdx == 23 || rowIdx == 25 || rowIdx == 31 || rowIdx == 37)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 20)
                        prefixName = "Crude Oil ";
                    if (rowIdx >= 27 && rowIdx <= 30)
                        prefixName = "Crude Oil Production Capacity ";
                    if (rowIdx >= 33 && rowIdx <= 36)
                        prefixName = "Surplus Crude Oil Production Capacity ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 38) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown){Console.WriteLine(petroleumDatum.Name);shown = true;}
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();            
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }                            
            Console.ReadKey();
            */

            #endregion  Sheet 3ctab

            #region Sheet 3dtab
            petroleumSheet = xlsx.Worksheet("3dtab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 43;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 10 || rowIdx == 13 || rowIdx == 15 || rowIdx == 18 || rowIdx == 20
                        || rowIdx == 25 || rowIdx == 27 || rowIdx == 30 || rowIdx == 32 || rowIdx == 40)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 34 && rowIdx <= 39)
                        prefixName = "World Real Gross Domestic Product (a) ";
                    if (rowIdx >= 42 && rowIdx <= 43)
                        prefixName = "Real U.S. Dollar Exchange Rate (a) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 42) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown){Console.WriteLine(petroleumDatum.Name);shown = true;}
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 3dtab

            #region Sheet 4atab
            petroleumSheet = xlsx.Worksheet("4atab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 7;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 62;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 16 || rowIdx == 34 || rowIdx == 35 || rowIdx == 45 || rowIdx == 47
                        || rowIdx == 49 )
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 7 && rowIdx <= 15)
                        prefixName = "Supply (million barrels per day) Crude Oil Supply ";
                    if (rowIdx >= 17 && rowIdx <= 33)
                        prefixName = "Supply (million barrels per day)  Other Supply ";
                    if (rowIdx >= 36 && rowIdx <= 44)
                        prefixName = "Consumption (million barrels per day) ";
                    if (rowIdx >= 50 && rowIdx <= 62)
                        prefixName = "End-of-period Inventories (million barrels) Commercial Inventory ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 62) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown){Console.WriteLine(petroleumDatum.Name);shown = true;}
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 4atab

            #region Sheet 4btab
            petroleumSheet = xlsx.Worksheet("4btab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 7;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 62;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 11 || rowIdx == 15 || rowIdx == 17 || rowIdx == 18 || rowIdx == 23
                        || rowIdx == 24 || rowIdx == 27 || rowIdx == 28 || rowIdx == 33 || rowIdx == 34
                        || rowIdx == 39 || rowIdx == 40 || rowIdx == 48 || rowIdx == 50 || rowIdx == 51
                        || rowIdx == 59)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 7 && rowIdx <= 10)
                        prefixName = "HGL Production Natural Gas Processing Plants";
                    if (rowIdx >= 12 && rowIdx <= 14)
                        prefixName = "HGL Production Refinery and Blender Net Production";
                    if (rowIdx == 16)
                        prefixName = "HGL Production Renewable Fuels and Oxygenate Plant Net Production";
                    if (rowIdx >= 19 && rowIdx <= 22)
                        prefixName = "HGL Net Imports ";
                    if (rowIdx >= 25 && rowIdx <= 26)
                        prefixName = "HGL Refinery and Blender Net Inputs ";
                    if (rowIdx >= 29 && rowIdx <= 32)
                        prefixName = "HGL Consumption ";
                    if (rowIdx >= 35 && rowIdx <= 38)
                        prefixName = "HGL Inventories (million barrels) ";
                    if (rowIdx >= 41 && rowIdx <= 47)
                        prefixName = "Refinery and Blender Net Inputs ";
                    if (rowIdx >= 52 && rowIdx <= 58)
                        prefixName = "Refinery and Blender Net Production ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 62) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown){Console.WriteLine(petroleumDatum.Name);shown = true;}
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();           
            */
            #endregion  Sheet 4btab

            #region Sheet 4ctab
            petroleumSheet = xlsx.Worksheet("4ctab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 27;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 7 || rowIdx == 15 || rowIdx == 16 || rowIdx == 17 || rowIdx == 24
                        || rowIdx == 26)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx == 6 || rowIdx == 14)
                        prefixName = "Prices (cents per gallon) ";
                    if (rowIdx >= 8 && rowIdx <= 13)
                        prefixName = "Prices (cents per gallon) Gasoline Regular Grade Retail Prices Including Taxes ";
                    if (rowIdx >= 18 && rowIdx <= 23)
                        prefixName = "End-of-period Inventories (million barrels) Total Gasoline Inventories ";
                    if (rowIdx == 25)
                        prefixName = "End-of-period Inventories (million barrels) Finished Gasoline Inventories ";
                    if (rowIdx == 27)
                        prefixName = "End-of-period Inventories (million barrels) Gasoline Blending Components Inventories ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 27) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 4ctab

            #region Sheet 5atab
            petroleumSheet = xlsx.Worksheet("5atab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 38;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 20 || rowIdx == 30)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 19)
                        prefixName = "Supply (billion cubic feet per day) ";
                    if (rowIdx >= 22 && rowIdx <= 29)
                        prefixName = "Consumption (billion cubic feet per day) ";
                    if (rowIdx >= 32 && rowIdx <= 38)
                        prefixName = "End-of-period Inventories (billion cubic feet) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 38) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey(); 
           */
            #endregion  Sheet 5atab

            #region Sheet 5btab
            petroleumSheet = xlsx.Worksheet("5btab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 39;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    //no empty rows to skip.
                    //if (rowIdx == x)
                    //break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx == 6)
                        prefixName = "Wholesale/Spot ";
                    if (rowIdx >= 8 && rowIdx <= 17)
                        prefixName = "Residential Retail ";
                    if (rowIdx >= 19 && rowIdx <= 28)
                        prefixName = "Commercial Retail ";
                    if (rowIdx >= 30 && rowIdx <= 39)
                        prefixName = "Industrial Retail ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 39) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey(); 
           */
            #endregion  Sheet 5atab

            #region Sheet 6tab
            petroleumSheet = xlsx.Worksheet("6tab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 45;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 16 || rowIdx == 20 || rowIdx == 28 || rowIdx == 30 || rowIdx == 38)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 15 )
                        prefixName = "Supply (million short tons) ";
                    if (rowIdx >= 17 && rowIdx <= 19)
                        prefixName = "Supply (million short tons) ";
                    if (rowIdx >= 22 && rowIdx <= 27)
                        prefixName = "Consumption (million short tons) ";
                    if (rowIdx == 29)
                        prefixName = "Consumption (million short tons) ";
                    if (rowIdx >= 32 && rowIdx <= 37)
                        prefixName = "End-of-period Inventories (million short tons) ";
                    if (rowIdx >= 41 && rowIdx <= 45)
                        prefixName = "Coal Market Indicators ";

                    string extraPrefix = "";
                    if(rowIdx == 41 )
                    { 
                        extraPrefix = "Coal Miner Productivity ";
                        prefixName = prefixName + extraPrefix;
                    }
                    if (rowIdx == 43)
                    {
                        extraPrefix = "Total Raw Steel Production ";
                        prefixName = prefixName + extraPrefix;
                    }
                    if (rowIdx == 45)
                    {
                        extraPrefix = "Cost of Coal to Electric Utilities ";
                        prefixName = prefixName + extraPrefix;
                    }


                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 45) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    //if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey(); 
           */
            #endregion  Sheet 6tab

            #region Sheet 7atab
            petroleumSheet = xlsx.Worksheet("7atab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 38;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    if (rowIdx == 12 || rowIdx == 28)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 11)
                        prefixName = "Electricity Supply (billion kilowatthours per day) ";
                    if (rowIdx >= 14 && rowIdx <= 22)
                        prefixName = "Electricity Consumption (billion kilowatthours per day) ";
                    if (rowIdx >= 25 && rowIdx <= 27)
                        prefixName = "End-of-period Fuel Inventories Held by Electric Power Sector ";
                    if (rowIdx >= 31 && rowIdx <= 34)
                        prefixName = "Prices Power Generation Fuel Costs (dollars per million Btu) ";
                    if (rowIdx >= 36 && rowIdx <= 38)
                        prefixName = "Prices Retail Prices (cents per kilowatthour) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 38) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                    year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey(); 
            */
            #endregion  Sheet 7atab

            #region Sheet 7btab
            petroleumSheet = xlsx.Worksheet("7btab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 52;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    //No empty rows
                    //if (rowIdx == 12 || rowIdx == 28)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 16)
                        prefixName = "Residential Sector ";
                    if (rowIdx >= 18 && rowIdx <= 28)
                        prefixName = "Commercial Sector ";
                    if (rowIdx >= 30 && rowIdx <= 40)
                        prefixName = "Industrial Sector ";
                    if (rowIdx >= 42 && rowIdx <= 52)
                        prefixName = "Total All Sectors (a) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 52) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 7btab

            #region Sheet 7ctab
            petroleumSheet = xlsx.Worksheet("7ctab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 48;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    //No empty rows
                    //if (rowIdx == 12 || rowIdx == 28)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 15)
                        prefixName = "Residential Sector ";
                    if (rowIdx >= 17 && rowIdx <= 26)
                        prefixName = "Commercial Sector ";
                    if (rowIdx >= 28 && rowIdx <= 37)
                        prefixName = "Industrial Sector ";
                    if (rowIdx >= 39 && rowIdx <= 48)
                        prefixName = "Total All Sectors (a) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 48) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 7btab

            #region Sheet 7dtab
            petroleumSheet = xlsx.Worksheet("7dtab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 60;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {
                    //No empty rows
                    //if (rowIdx == 12 || rowIdx == 28)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 20)
                        prefixName = "United States ";
                    if (rowIdx >= 22 && rowIdx <= 30)
                        prefixName = "Northeast Census Region ";
                    if (rowIdx >= 32 && rowIdx <= 40)
                        prefixName = "South Census Region ";
                    if (rowIdx >= 42 && rowIdx <= 50)
                        prefixName = "Midwest Census Region ";
                    if (rowIdx >= 50 && rowIdx <= 60)
                        prefixName = "West Census Region ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 60) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 7dtab

            #region Sheet 7etab
            petroleumSheet = xlsx.Worksheet("7etab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 7;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 35;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (rowIdx == 30)
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string extraPrefix = "Fuel Consumption for Electricity Generation, All Sectors ";
                    string prefixName = "";
                    if (rowIdx >= 7 && rowIdx <= 13)
                        prefixName = extraPrefix+"United States ";
                    if (rowIdx >= 15 && rowIdx <= 17)
                        prefixName = extraPrefix + "Northeast Census Region ";
                    if (rowIdx >= 19 && rowIdx <= 21)
                        prefixName = extraPrefix + "South Census Region ";
                    if (rowIdx >= 23 && rowIdx <= 25)
                        prefixName = extraPrefix + "Midwest Census Region ";
                    if (rowIdx >= 27 && rowIdx <= 29)
                        prefixName = extraPrefix + "West Census Region ";
                    if (rowIdx >= 32 && rowIdx <= 35)
                        prefixName =  "End-of-period U.S. Fuel Inventories Held by Electric Power ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 35) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 7etab

            #region Sheet 8tab
            petroleumSheet = xlsx.Worksheet("8tab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 44;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    //No empty rows
                    //if (rowIdx == 30)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 12)
                        prefixName =  "Electric Power Sector  ";
                    if (rowIdx >= 14 && rowIdx <= 19)
                        prefixName =  "Industrial Sector ";
                    if (rowIdx >= 21 && rowIdx <= 24)
                        prefixName =  "Commercial Sector ";
                    if (rowIdx >= 26 && rowIdx <= 29)
                        prefixName =  "Residential Sector ";
                    if (rowIdx >= 31 && rowIdx <= 33)
                        prefixName =  "Transportation Sector  ";
                    if (rowIdx >= 35 && rowIdx <= 44)
                        prefixName = "All Sectors Total ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 44) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 8tab

            #region Sheet 9btab
            petroleumSheet = xlsx.Worksheet("9btab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 54;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    //No empty rows
                    //if (rowIdx == 30)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 14)
                        prefixName = "Real Gross State Product (Billion $2009) ";
                    if (rowIdx >= 16 && rowIdx <= 24)
                        prefixName = "Industrial Output, Manufacturing (Index, Year 2012=100) ";
                    if (rowIdx >= 26 && rowIdx <= 34)
                        prefixName = "Real Personal Income (Billion $2009) ";
                    if (rowIdx >= 36 && rowIdx <= 44)
                        prefixName = "Households (Thousands) ";
                    if (rowIdx >= 46 && rowIdx <= 54)
                        prefixName = "Total Non-farm Employment (Millions) ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 54) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 9btab

            #region Sheet 9ctab
            petroleumSheet = xlsx.Worksheet("9ctab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 6;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 48;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    //No empty rows
                    //if (rowIdx == 30)
                    //    break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";

                    if (rowIdx >= 6 && rowIdx <= 15)
                        prefixName = "Heating Degree Days ";
                    if (rowIdx >= 17 && rowIdx <= 26)
                        prefixName = "Heating Degree Days, Prior 10-year Average ";
                    if (rowIdx >= 28 && rowIdx <= 37)
                        prefixName = "Cooling Degree Days ";
                    if (rowIdx >= 39 && rowIdx <= 48)
                        prefixName = "Cooling Degree Days, Prior 10-year Average ";

                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 48) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 9ctab

            #region Sheet 9atab
            petroleumSheet = xlsx.Worksheet("9atab");
            petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            firstRowWithData = 7;
            cellShift = 3;
            rowCellHeaders = 4;
            lastCellWithData = 74;
            lastRow = 69;
            category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                var previusRow = petroleumSheet.Row(rowIdx-1);
                int cellCnt = 1;
                int year = 2012;
                Boolean shown = false;
                for (int cellIdx = cellShift; cellIdx <= lastCellWithData; cellIdx++)
                {

                    if (rowIdx == 28 || rowIdx == 42 || rowIdx == 52 || rowIdx == 64 || string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString()))
                        break;

                    GovUSAShortTermEnergyOutlook petroleumDatum = new GovUSAShortTermEnergyOutlook();
                    petroleumDatum.Source = URL;

                    if (string.IsNullOrEmpty(petroleumRow.Cell(1).Value.ToString().Trim()))
                    {
                        continue;
                    }

                    string prefixName = "";
                    if (rowIdx >= 6 && rowIdx <= 27)
                        prefixName = "Macroeconomic ";
                    if (rowIdx >= 30 && rowIdx <= 41)
                        prefixName = "Industrial Production Indices (Index, 2012=100) ";
                    if (rowIdx >= 44 && rowIdx <= 51)
                        prefixName = "Price Indexes ";
                    if (rowIdx >= 54 && rowIdx <= 63)
                        prefixName = "Miscellaneous ";
                    if (rowIdx >= 66 && rowIdx <= 69)
                        prefixName = "Carbon Dioxide (CO2) Emissions (million metric tons) ";
                    if (string.IsNullOrEmpty(previusRow.Cell(1).Value.ToString()))
                    {
                        petroleumDatum.Name = prefixName + previusRow.Cell(2).Value.ToString() + " " + petroleumRow.Cell(2).Value.ToString();
                    }
                    else
                    {
                        petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();
                    }
                    petroleumDatum.Code = petroleumRow.Cell(1).Value.ToString();
                    
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();

                    GetStartEndDate(year, petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString(), out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 69) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);
                    if (cellCnt % 12 == 0)
                        year++;

                    cellCnt++;

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=7)break;
            }
            //Console.ReadKey();    
            /*
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();
            */
            #endregion  Sheet 9atab

        }

        private void GetStartEndDate(int year, string month, out DateTime StartDate, out DateTime EndDate)
        {
            DateTime StartDt = new DateTime(), EndDt = new DateTime();

            if (month.Contains("Jan"))
            {
                StartDt = new DateTime(int.Parse(year.ToString().ToString()), 1, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 1, 31);
            }
            else if (month.Contains("Feb"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 2, 1);
                EndDt = DateTime.IsLeapYear(int.Parse(year.ToString())) ? new DateTime(int.Parse(year.ToString()), 2, 29) : new DateTime(int.Parse(year.ToString()), 2, 28);
            }
            else if (month.Contains("Mar"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 3, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 3, 31);
            }
            else if (month.Contains("Apr"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 4, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 4, 30);
            }
            else if (month.Contains("May"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 5, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 5, 31);
            }
            else if (month.Contains("Jun"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 6, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 6, 30);
            }
            else if (month.Contains("Jul"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 7, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 7, 31);
            }
            else if (month.Contains("Aug"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 8, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 8, 31);
            }
            else if (month.Contains("Sep"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 9, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 9, 30);
            }
            else if (month.Contains("Oct"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 10, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 10, 31);
            }
            else if (month.Contains("Nov"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 11, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 11, 30);
            }
            else if (month.Contains("Dec"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 12, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 12, 31);
            }

            StartDate = StartDt;
            EndDate = EndDt;
        }

    }
}