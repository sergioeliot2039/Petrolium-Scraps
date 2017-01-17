using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using System.IO;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace ScrappingDataFromWebPageInCS
{
    class Program
    {
        static void Main(string[] args)
        {

            /*string[] xlsFilePaths =
            {
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.1.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.2.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.3.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.4.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.5.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.6.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.7.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.10.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.11.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.12.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\ET_3.13.xls",
                @"C:\Users\sergi_000\Desktop\NOV006\JODI_October_2016.xls"
            };*/


            foreach (string xlsFilePath in args)
            {
                GetExcelSheetData(xlsFilePath);
                Console.WriteLine("Click enter to scrap from next file or exit if it is the last...");
                Console.ReadLine();
                
            }
            


        }

        static string GetExcelSheetData(string urlOfExcelFile)
        {
            var ios = new FileStream(urlOfExcelFile,FileMode.Open);
            var a = new POIFSFileSystem(ios);
            
            var workbook = new HSSFWorkbook(a);
            HSSFSheet worksheet;

            string fileName = Path.GetFileNameWithoutExtension(urlOfExcelFile);
            string source = "";
            switch(fileName)
            {
                case "ET_3.1": 
                    source = "ET_3.1";
                    List<SupplyOfProductsByPeriod> supplyData = GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "Indigenous Production", string.Empty, 1, source);
                    supplyData.AddRange(GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "Indigenous Production", string.Empty, 1, source));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach(SupplyOfProductsByPeriod s in supplyData)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.Year, s.Quarter);
                    }
                    break;

                case "ET_3.2":
                    source = "ET_3.2";
                    supplyData = GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "Indigenous Production", string.Empty, 1, source);
                    supplyData.AddRange(GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "Indigenous Production", string.Empty, 1, source));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach(SupplyOfProductsByPeriod s in supplyData)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.Year, s.Quarter);
                    }
                    break;
                case "ET_3.3": 
                    source = "ET_3.3";
                    List<SupplyOfPetProductsByPeriod> petProdSupplyData = GetSupplyOfPetProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "Indigenous Production",source);
                    petProdSupplyData.AddRange(GetSupplyOfPetProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "Indigenous Production",source));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach (SupplyOfPetProductsByPeriod s in petProdSupplyData)
                    {
                        //Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} \t {5} \t {6} \t {7} \t {8} \t {9} \t {10} \t {11} \t {12} \t {13}"
                                             
                        Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} \t {5} \t {6} \t {7} \t {8} \t {9} \t {10} \t {11}"
                            , s.Name, s.Year, s.Quarter, s.TotalPetroleumProducts, s.MotorSpirit
                            , s.DERV, s.GasOil, s.AviationTurbineFuel, s.FuelOils, s.PetroleumGases, s.BurningOil
                            , s.OtherProducts);
                    }
                    break;
                case "ET_3.4": 
                    source = "ET_3.4";
                    petProdSupplyData = GetSupplyOfPetProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "Indigenous Production",source);
                    petProdSupplyData.AddRange(GetSupplyOfPetProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "Indigenous Production",source));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach (SupplyOfPetProductsByPeriod s in petProdSupplyData)
                    {
                        //Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} \t {5} \t {6} \t {7} \t {8} \t {9} \t {10} \t {11} \t {12} \t {13}"
                                             
                        Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} \t {5} \t {6} \t {7} \t {8} \t {9} \t {10} \t {11}"
                            , s.Name, s.Year, s.Quarter, s.TotalPetroleumProducts, s.MotorSpirit
                            , s.DERV, s.GasOil, s.AviationTurbineFuel, s.FuelOils, s.PetroleumGases, s.BurningOil
                            , s.OtherProducts);
                    }
                    break;
                case "ET_3.5":
                    List<SupplyOfProductsByPeriod> demandData = GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "Motor spirit", "Heavy", 2, "ET_3.5");
                    demandData.AddRange(GetSupplyOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "Motor spirit", "Heavy", 2, "ET_3.5"));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach (SupplyOfProductsByPeriod s in demandData)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Name, s.Quantity, s.Year, s.Quarter);
                    }
                    break;
                case "ET_3.6":
                    List<SupplyOfProductsByPeriod> stockData = GetStockOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1997", 2, 3, "ET_3.6");
                    stockData.AddRange(GetStockOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1995",3,3, "ET_3.6"));
                    Console.WriteLine("Concept \t\t Quantity \t Year \t Quarter");
                    foreach (SupplyOfProductsByPeriod s in stockData)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3}", s.Year, s.Quarter,s.Name, s.Quantity );
                    }
                    break;
                case "ET_3.7":
                    List<SupplyOfProductsByPeriod> drillingData = GetDrillingActivityPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1998", 2, "ET_3.7");
                    drillingData.AddRange(GetDrillingActivityPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1998", 2, "ET_3.7"));                    
                    Console.WriteLine("Concept \t Year \t Quarter \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in drillingData)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} ", s.Name, s.Quarter, s.Year, s.Month, s.Quantity, s.Source);
                    }
                    break;
                case "ET_3.10":
                    List<SupplyOfProductsByPeriod> importExportData = GetImportExportDataPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1995", 2, "ET_3.10");
                    importExportData.AddRange(GetImportExportDataPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1995", 2, "ET_3.10"));
                    importExportData.AddRange(GetImportExportDataPeriod(workbook.GetSheet(TargetSheets.MONTH.ToString()), "1995", 2, "ET_3.10"));
                    Console.WriteLine("Concept \t Year \t Quarter \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in importExportData)
                    {
                        Console.WriteLine("{0} \t {1} {2} {3} \t {4} \t {5} \t {6}", s.Name, s.Quarter, s.Year, s.Month, s.Quantity,s.StartDate.ToShortDateString(),s.EndDate.ToShortDateString());
                        //Console.WriteLine("{0} \t {1} {2} ", s.Quarter, s.StartDate.ToShortDateString(), s.EndDate.ToShortDateString());
                    }
                    break;
                case "ET_3.11":
                    List<SupplyOfProductsByPeriod> stockData2 = GetStockOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1995", 2, 3, "ET_3.11");
                    stockData2.AddRange(GetStockOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1995", 2, 3, "ET_3.11"));
                    stockData2.AddRange(GetStockOfProductsByPeriod(workbook.GetSheet(TargetSheets.MONTH.ToString()), "1995", 2, 3, "ET_3.11"));
                    Console.WriteLine("Concept \t\t Quarter  \t Year \t Month \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in stockData2)
                    {
                        Console.WriteLine("{0} \t {1} \t {2} \t {3} \t {4} \t {5}", s.Name, s.Quarter, s.Year, s.Month, s.Quantity, s.Source);
                        
                    }
                    break;
                case "ET_3.12":
                    List<SupplyOfProductsByPeriod> throughputData = GetThroughtputOfProductsByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1995", 1, "ET_3.12");
                    throughputData.AddRange(GetThroughtputOfProductsByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1995", 2, "ET_3.12"));
                    throughputData.AddRange(GetThroughtputOfProductsByPeriod(workbook.GetSheet(TargetSheets.MONTH.ToString()), "1995", 2, "ET_3.12"));
                    Console.WriteLine("Concept \t\t Quarter  \t Year \t Month \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in throughputData)
                    {
                        Console.WriteLine("{0} \t {1} {2} {3} \t {4} \t {5} \t {6}", s.Name, s.Quarter, s.Year, s.Month, s.Quantity, s.StartDate.ToShortDateString(), s.EndDate.ToShortDateString());                        
                    }
                    break;
                case "ET_3.13":
                    List<SupplyOfProductsByPeriod> inlandConsumptionData = GetInlandConsumptionByPeriod(workbook.GetSheet(TargetSheets.ANUAL.ToString()), "1998", 2, "ET_3.13");
                    inlandConsumptionData.AddRange(GetInlandConsumptionByPeriod(workbook.GetSheet(TargetSheets.QUARTER.ToString()), "1998", 2, "ET_3.13"));
                    inlandConsumptionData.AddRange(GetInlandConsumptionByPeriod(workbook.GetSheet(TargetSheets.MONTH.ToString()), "1998", 2, "ET_3.13"));
                    Console.WriteLine("Concept \t\t Quarter  \t Year \t Month \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in inlandConsumptionData)
                    {
                        Console.WriteLine("{0} \t {1} {2} {3} \t {4} \t {5} \t {6}", s.Name, s.Quarter, s.Year, s.Month, s.Quantity, s.StartDate.ToShortDateString(), s.EndDate.ToShortDateString());
                    }
                    break;
                case "JODI_October_2016":
                    List<SupplyOfProductsByPeriod> crudOildata = GrabCrudOilTable(workbook.GetSheet(TargetSheets.MONTH1.ToString()), " Production", 2, "JODI_October_2016");                     //Notice space on " Production"..
                    crudOildata.AddRange(GrabPetProdsTable(workbook.GetSheet(TargetSheets.MONTH1.ToString()), " Production", 6, "JODI_October_2016"));
                    crudOildata.AddRange(GrabCrudOilTable(workbook.GetSheet(TargetSheets.MONTH2.ToString()), " Production", 2, "JODI_October_2016"));
                    crudOildata.AddRange(GrabPetProdsTable(workbook.GetSheet(TargetSheets.MONTH2.ToString()), " Production", 6, "JODI_October_2016"));
                    Console.WriteLine("Name \t Quantity \t Source ");
                    foreach (SupplyOfProductsByPeriod s in crudOildata)
                    {
                        Console.WriteLine("{0} \t {1} {2} {3} \t {4} \t {5} \t {6}", s.Name, s.Quarter, s.Year, s.Month, s.Quantity, s.StartDate.ToShortDateString(), s.EndDate.ToShortDateString());
                    }
                    break;
                default:
                    break;
            }
            


            //var worksheet= workbook.GetSheet():
            return "";
        }

        private static List<SupplyOfProductsByPeriod> GrabCrudOilTable(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);

            for (int rowIdx = firstrow; rowIdx < 17; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx <= worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx++)
                {

                    SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                    stockOfProduct.Source = source;
                    stockOfProduct.Quarter = "-";
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx).CellType == CellType.Numeric)
                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;
                    if (rowIdx == 14 || rowIdx == 15)
                        stockOfProduct.Name = "Crud Oil - Stocks - " + worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                    else
                        stockOfProduct.Name = "Crud Oil - " + worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue;
                    supOfProdsByPerLst.Add(stockOfProduct);




                }

            }

            return supOfProdsByPerLst;     
            
        }

        //JODI_October_2016 Sheets Month-1 & Month-2 table with Crud Oil column.
        private static List<SupplyOfProductsByPeriod> GrabPetProdsTable(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
           List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();            
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);            

            for (int rowIdx = firstrow; rowIdx < 17; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
if(worksheet.GetRow(firstrow - 1).GetCell(cellIdx)!= null)
{ 
                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Quarter = "-";
                        if (worksheet.GetRow(rowIdx).GetCell(cellIdx).CellType==CellType.Numeric)
                            stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;
                        if (rowIdx == 15 || rowIdx == 14)
                            stockOfProduct.Name = " Stocks "+ worksheet.GetRow(rowIdx).GetCell(6).StringCellValue + " - Petroleum Products - " + worksheet.GetRow(firstrow - 1).GetCell(cellIdx).StringCellValue;
                        else
                            stockOfProduct.Name = worksheet.GetRow(rowIdx).GetCell(5).StringCellValue + " -Petroleum Products - " + worksheet.GetRow(firstrow - 1).GetCell(cellIdx).StringCellValue;

                        supOfProdsByPerLst.Add(stockOfProduct);

}
                    

                }

            }

            return supOfProdsByPerLst;        	        
        }

        //ET_3.13 Row scan with columns for document ET_3.13
        private static List<SupplyOfProductsByPeriod> GetInlandConsumptionByPeriod(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = double.Parse(dataAtFirstRowFirstCell);//It is assumed year info is at first cell of first row with data.
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            Console.WriteLine("Frist row is {0}", firstrow);

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx) != null && worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric)
                    {
                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Year = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                        {

                            Regex regex = new Regex("[1-4]");
                            MatchCollection m = regex.Matches(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            {
                                qinfo = ma.Value;
                                break;
                            }

                            stockOfProduct.Quarter = "Q" + qinfo;
                        }
                        else
                        {
                            stockOfProduct.Quarter = "-";

                        }
                        if (worksheet.SheetName.Equals(TargetSheets.MONTH.ToString()))
                        {
                            stockOfProduct.Month = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                        }
                        else
                        {
                            stockOfProduct.Month = "-";
                        }
                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;

                        switch (cellIdx)
                        {
                            case 2:
                                stockOfProduct.Name = "Total";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 3:
                                stockOfProduct.Name = "Butane and propane";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 4:
                                stockOfProduct.Name = "Other Petroleum Gases";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 5:
                                stockOfProduct.Name = "Naptha (LDF)";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 6:
                                stockOfProduct.Name = "Motor spirit";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 7:
                                stockOfProduct.Name = "Kerosene - Aviation turbine fuel";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 8:
                                stockOfProduct.Name = "Kerosene - Burning oil";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 9:
                                stockOfProduct.Name = "Gas/diesel oil - Derv fuel";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 10:
                                stockOfProduct.Name = "Gas/diesel oil - Other";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 11:
                                stockOfProduct.Name = "Fuel oil";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 12:
                                stockOfProduct.Name = "Lubricating oils";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 13:
                                stockOfProduct.Name = "Bitumen";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                        }

                    }

                }

            }

            return supOfProdsByPerLst;       
        }

        //ET_3.12 Row scan with columns for document ET_3.12
        private static List<SupplyOfProductsByPeriod> GetThroughtputOfProductsByPeriod(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = double.Parse(dataAtFirstRowFirstCell);//It is assumed year info is at first cell of first row with data.
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            Console.WriteLine("Frist row is {0}", firstrow);

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx) != null && worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric)
                    {
                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Year = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                        {

                            Regex regex = new Regex("[1-4]");
                            MatchCollection m = regex.Matches(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            {
                                qinfo = ma.Value;
                                break;
                            }

                            stockOfProduct.Quarter = "Q" + qinfo;
                        }
                        else
                        {
                            stockOfProduct.Quarter = "-";

                        }
                        if (worksheet.SheetName.Equals(TargetSheets.MONTH.ToString()))
                        {
                            stockOfProduct.Month = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                        }
                        else
                        {
                            stockOfProduct.Month = "-";
                        }
                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()))
                        {
                            switch (cellIdx)
                            {
                                case 1:
                                    stockOfProduct.Name = "Throughput of crude and process oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 2:
                                    stockOfProduct.Name = "Refinery use - Fuel";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 3:
                                    stockOfProduct.Name = "Refinery use - Losses/(gains)";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 4:
                                    stockOfProduct.Name = "Total output of petroleum products";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 5:
                                    stockOfProduct.Name = "Gases - Butane and propane";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 6:
                                    stockOfProduct.Name = "Gases - Other petroleum";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 7:
                                    stockOfProduct.Name = "Naphtha (LDF)";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 8:
                                    stockOfProduct.Name = "Motor spirit";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 9:
                                    stockOfProduct.Name = "Kerosene - Aviation turbine fuel";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 10:
                                    stockOfProduct.Name = "Kerosene - Burning oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 11:
                                    stockOfProduct.Name = "Gas oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 12:
                                    stockOfProduct.Name = "DERV oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 13:
                                    stockOfProduct.Name = "Fuel oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 14:
                                    stockOfProduct.Name = "Lubricating oils";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 15:
                                    stockOfProduct.Name = "Bitumen";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                            }
                        }
                        else
                        { 
                            switch (cellIdx)
                            {
                                case 2:
                                    stockOfProduct.Name = "Throughput of crude and process oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 3:
                                    stockOfProduct.Name = "Refinery use - Fuel";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 4:
                                    stockOfProduct.Name = "Refinery use - Losses/(gains)";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 5:
                                    stockOfProduct.Name = "Total output of petroleum products";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 6:
                                    stockOfProduct.Name = "Gases - Butane and propane";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 7:
                                    stockOfProduct.Name = "Gases - Other petroleum";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 8:
                                    stockOfProduct.Name = "Naphtha (LDF)";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 9:
                                    stockOfProduct.Name = "Motor spirit";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 10:
                                    stockOfProduct.Name = "Kerosene - Aviation turbine fuel";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 11:
                                    stockOfProduct.Name = "Kerosene - Burning oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 12:
                                    stockOfProduct.Name = "Gas oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 13:
                                    stockOfProduct.Name = "DERV oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 14:
                                    stockOfProduct.Name = "Fuel oil";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 15:
                                    stockOfProduct.Name = "Lubricating oils";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                                case 16:
                                    stockOfProduct.Name = "Bitumen";
                                    supOfProdsByPerLst.Add(stockOfProduct);
                                    break;
                            }
                        }

                    }

                }

            }

            return supOfProdsByPerLst;             
        }


        //ET_3.10 Row scan with columns for document ET_3.10
        private static List<SupplyOfProductsByPeriod> GetImportExportDataPeriod(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = double.Parse(dataAtFirstRowFirstCell);//It is assumed year info is at first cell of first row with data.
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            Console.WriteLine("Frist row is {0}", firstrow);

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx) != null && worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric)
                    {
                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Year = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                        {

                            Regex regex = new Regex("[1-4]");
                            MatchCollection m = regex.Matches(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            {
                                qinfo = ma.Value;
                                break;
                            }

                            stockOfProduct.Quarter = "Q" + qinfo;
                        }
                        else
                        {
                            stockOfProduct.Quarter = "-";

                        }
                        if (worksheet.SheetName.Equals(TargetSheets.MONTH.ToString()))
                        {
                            stockOfProduct.Month = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                        }
                        else
                        {
                            stockOfProduct.Month = "-";
                        }
                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;

                        switch (cellIdx)
                        {
                            case 2:
                                stockOfProduct.Name = "Indigenous production - Total";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 3:
                                stockOfProduct.Name = "Indigenous production - Crude oil";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 4:
                                stockOfProduct.Name = "Indigenous production - NGLs";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 5:
                                stockOfProduct.Name = "Indigenous production - Feedstocks";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 6:
                                stockOfProduct.Name = "Refinery receipts - Total receipts";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 7:
                                stockOfProduct.Name = "Refinery receipts - Indigenous";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 8:
                                stockOfProduct.Name = "Foreign Trade - Net Imports/exports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 9:
                                stockOfProduct.Name = "Foreign Trade - Crude oil and NGLs - Imports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 10:
                                stockOfProduct.Name = "Foreign Trade - Crude oil and NGLs - Exports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 11:
                                stockOfProduct.Name = "Foreign Trade - Process oils - Imports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 12:
                                stockOfProduct.Name = "Foreign Trade - Process oils - Exports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 13:
                                stockOfProduct.Name = "Foreign Trade - Petroleum products - Imports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 14:
                                stockOfProduct.Name = "Foreign Trade - Petroleum products - Exports";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 15:
                                stockOfProduct.Name = "Foreign Trade - Petroleum products - Bunkers";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                        }

                    }

                }

            }

            return supOfProdsByPerLst;            
        }

        private static List<SupplyOfProductsByPeriod> GetDrillingActivityPeriod(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = double.Parse(dataAtFirstRowFirstCell);//It is assumed year info is at first cell of first row with data.
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            Console.WriteLine("Frist row is {0}", firstrow);

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {


                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx) != null && worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric)
                    {



                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Year = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                        {

                            Regex regex = new Regex("[1-4]");
                            MatchCollection m = regex.Matches(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            {
                                qinfo = ma.Value;
                                break;
                            }

                            stockOfProduct.Quarter = "Q" + qinfo;
                        }
                        else
                        {
                            stockOfProduct.Quarter = "-";

                        }
                        if (worksheet.SheetName.Equals(TargetSheets.MONTH.ToString()))
                        {
                            stockOfProduct.Month = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                        }
                        else
                        {
                            stockOfProduct.Month = "-";
                        }
                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;
                        switch(cellIdx)
                        {
                            case 2:
                                stockOfProduct.Name = "Offshore - Exploration";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 3:
                                stockOfProduct.Name = "Offshore - Appraisal";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 4:
                                stockOfProduct.Name = "Offshore - Exploration & appraisal";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 5:
                                stockOfProduct.Name = "Offshore - Development";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 6:
                                stockOfProduct.Name = "Onshore - Exploration & appraisal";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;
                            case 7:
                                stockOfProduct.Name = "Onshore - Development";
                                supOfProdsByPerLst.Add(stockOfProduct);
                                break;

                        }                                                 

                    }

                }

            }

            return supOfProdsByPerLst;
        }

        private static List<SupplyOfProductsByPeriod> GetStockOfProductsByPeriod(ISheet worksheet, string dataAtFirstRowFirstCell, int cellShift, int rowUpperCatShift, string source)
        {
            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = double.Parse(dataAtFirstRowFirstCell);//It is assumed year info is at first cell of first row with data.
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            Console.WriteLine("Frist row is {0}",firstrow);

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {

                string upperCategory = "";
                
                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx) != null && worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric)
                    {



                        SupplyOfProductsByPeriod stockOfProduct = new SupplyOfProductsByPeriod();
                        stockOfProduct.Source = source;
                        stockOfProduct.Year = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue;

                        if (worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                        {

                            Regex regex = new Regex("[1-4]");
                            MatchCollection m = regex.Matches(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue);
                            string qinfo = "";
                            foreach (Match ma in m)
                            { 
                                qinfo = ma.Value;
                                break;
                            }

                            stockOfProduct.Quarter = "Q" + qinfo;
                        }
                        else
                        {
                            stockOfProduct.Quarter = "-";
                        
                        }
                        if (worksheet.SheetName.Equals(TargetSheets.MONTH.ToString()))
                        {
                            stockOfProduct.Month = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum + 1).StringCellValue;
                        }
                        else
                        {
                            stockOfProduct.Month = "-";
                        }

                        stockOfProduct.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;
                        string upperCategoryBakUP = upperCategory;

                        try
                        {
                            upperCategory = worksheet.GetRow(firstrow - rowUpperCatShift).GetCell(cellIdx).StringCellValue;
                            if (string.IsNullOrEmpty(upperCategory))
                            {
                                upperCategory = worksheet.GetRow(firstrow - rowUpperCatShift-1).GetCell(cellIdx).StringCellValue;
                            }
                            if (string.IsNullOrEmpty(upperCategory))
                            {
                                upperCategory = upperCategoryBakUP;
                            }


                            if ( worksheet.GetRow(firstrow - rowUpperCatShift + 2).GetCell(cellIdx) != null && worksheet.GetRow(firstrow - rowUpperCatShift + 2).GetCell(cellIdx).CellType == CellType.String)
                            {
                                stockOfProduct.Name = upperCategory + " - " + worksheet.GetRow(firstrow - rowUpperCatShift + 1).GetCell(cellIdx).StringCellValue
                                    + worksheet.GetRow(firstrow - rowUpperCatShift + 2).GetCell(cellIdx).StringCellValue;
                            }
                            else
                            {
                                stockOfProduct.Name = upperCategory + " - " + worksheet.GetRow(firstrow - rowUpperCatShift + 1).GetCell(cellIdx).StringCellValue;
                            }
                            supOfProdsByPerLst.Add(stockOfProduct);
                        }
                        catch { }
                    }

                }
                
            }

            return supOfProdsByPerLst;

        }

        // GetSupplyOfPetProductsByPeriod
        // Scraps data from files ET_3.3.xls & ET_3.4.xl
        // Data is returned as model class SupplyOfProductsByPeriod
        private static List<SupplyOfPetProductsByPeriod> GetSupplyOfPetProductsByPeriod(NPOI.SS.UserModel.ISheet worksheet, string dataAtFirstRowFirstCell, string source)
        {
            List<SupplyOfPetProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfPetProductsByPeriod>();
            double firstyear = double.Parse(GetFirstYearSuppPetProd(worksheet, dataAtFirstRowFirstCell).Trim().Substring(0,4));
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            double actualyear = firstyear;
            short quartercount = 1;

            int yearIncrementSize = worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()) ? 8 : 9;

            for (int rowIdx = firstrow; rowIdx < worksheet.LastRowNum; rowIdx++)
            {


                actualyear = firstyear;
                quartercount = 1;
                yearIncrementSize = worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()) ? 8 : 9;
                int cellPtr = 1;
                //Console.WriteLine(actualyear);
                //Console.WriteLine("Renglón: {0}",rowIdx);
                SupplyOfPetProductsByPeriod supOfProdsByPer = new SupplyOfPetProductsByPeriod();
                for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + 1; cellIdx < worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                {
                    //Console.WriteLine(cellPtr);
                    //Console.WriteLine(worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue);

                    supOfProdsByPer.Source = source;
                    supOfProdsByPer.Name = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue;
                    supOfProdsByPer.Year = actualyear;
                    supOfProdsByPer.Quarter = worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()) ? "-" : "Q" + quartercount;

                    if (worksheet.GetRow(rowIdx).GetCell(cellIdx)!=null)
                    {
                            try{
                                
                                if (cellPtr % yearIncrementSize == 1)
                                    supOfProdsByPer.TotalPetroleumProducts = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                if (cellPtr % yearIncrementSize == 2)
                                    supOfProdsByPer.MotorSpirit = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;

                                if (actualyear >= 2005 || worksheet.SheetName.Equals(TargetSheets.QUARTER.ToString()))
                                {
                                    if (cellPtr % yearIncrementSize == 3)
                                        supOfProdsByPer.DERV = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 4)
                                        supOfProdsByPer.GasOil = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 5)
                                        supOfProdsByPer.AviationTurbineFuel = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 6)
                                        supOfProdsByPer.FuelOils = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 7)
                                        supOfProdsByPer.PetroleumGases = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 8)
                                        supOfProdsByPer.BurningOil = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 0)
                                        supOfProdsByPer.OtherProducts = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;

                                }
                                else
                                {
                                    supOfProdsByPer.DERV = 0.0;
                                    if(cellPtr % yearIncrementSize == 3)
                                        supOfProdsByPer.GasOil = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 4)
                                        supOfProdsByPer.AviationTurbineFuel = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 5)
                                        supOfProdsByPer.FuelOils = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 6)
                                        supOfProdsByPer.PetroleumGases = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 7)
                                        supOfProdsByPer.BurningOil = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                    if (cellPtr % yearIncrementSize == 0)
                                        supOfProdsByPer.OtherProducts = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                                }

                            //supOfProdsByPer.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue != null ? worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue : 0.0;
                            
                            }
                            catch {  }


                            if (worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()))
                            {
                                if (cellPtr % yearIncrementSize == 0 && (cellIdx != worksheet.GetRow(rowIdx).FirstCellNum + 1))
                                {
                                    if (!supOfProdsByPer.Name.Contains("Return to") && !string.IsNullOrEmpty(supOfProdsByPer.Name))
                                    {
                                        supOfProdsByPerLst.Add(supOfProdsByPer);
                                        supOfProdsByPer = new SupplyOfPetProductsByPeriod();
                                    }
                                    actualyear = actualyear + 1;
                                    if (actualyear == 2005)
                                        yearIncrementSize = 9;
                                    //Console.WriteLine(actualyear);
                                    cellPtr = 1;
                                }
                                else
                                {
                                    cellPtr++;
                                }
                            }
                            else
                            {
                                
                                if (cellPtr % yearIncrementSize == 0 && (cellIdx != worksheet.GetRow(rowIdx).FirstCellNum + 1))
                                {
                                    if (!supOfProdsByPer.Name.Contains("Return to") && !string.IsNullOrEmpty(supOfProdsByPer.Name))
                                    {
                                        supOfProdsByPerLst.Add(supOfProdsByPer);
                                        supOfProdsByPer = new SupplyOfPetProductsByPeriod();
                                    }
                                    quartercount++;
                                    cellPtr = 1;
                                    if (quartercount >= 5)
                                    {
                                        quartercount = 1;
                                        actualyear = actualyear + 1;
                                        //Console.WriteLine(actualyear);
                                    }
                                    //Console.WriteLine("Q"+quartercount.ToString());
                                }
                                else
                                {
                                    cellPtr++;
                                }
                            }
                        
                    }
                    /*else
                    {
                        Console.WriteLine(rowIdx);
                        Console.WriteLine("Fyera");
                    }*/


                }



            }
            return supOfProdsByPerLst;
        }



        //GetSupplyOfProductsByPeriod
        // Scraps data from files ET_3.1.xls & ET_3.2.xl
        // Data is returned as model class SupplyOfProductsByPeriod
        private static List<SupplyOfProductsByPeriod> GetSupplyOfProductsByPeriod(NPOI.SS.UserModel.ISheet worksheet, string dataAtFirstRowFirstCell, string lastRowName, int cellShift, string source)
        {

            List<SupplyOfProductsByPeriod> supOfProdsByPerLst = new List<SupplyOfProductsByPeriod>();
            double firstyear = GetFirstYearSuppProd(worksheet, dataAtFirstRowFirstCell);
            int firstrow = GetFirstRowSuppProd(worksheet, dataAtFirstRowFirstCell);
            double actualyear = firstyear;
            short quartercount = 1;
            for(int rowIdx = firstrow;rowIdx<worksheet.LastRowNum;rowIdx++)
            {

                    
                    actualyear = firstyear;
                    quartercount = 1;

                    if (worksheet.GetRow(rowIdx)!=null)
                    { 
                        for (int cellIdx = worksheet.GetRow(rowIdx).FirstCellNum + cellShift; cellIdx<worksheet.GetRow(rowIdx).LastCellNum; cellIdx++)
                        {
                            SupplyOfProductsByPeriod supOfProdsByPer = new SupplyOfProductsByPeriod();
                            supOfProdsByPer.Source = source;
                            supOfProdsByPer.Name = worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue;
                            supOfProdsByPer.Year = actualyear;
                            supOfProdsByPer.Quarter = worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()) ? "-" : "Q" + quartercount;
                            if (worksheet.GetRow(rowIdx).GetCell(cellIdx)!=null)
                            { 
                                supOfProdsByPer.Quantity = worksheet.GetRow(rowIdx).GetCell(cellIdx).NumericCellValue;

                                if (!supOfProdsByPer.Name.Contains("Return to") && !string.IsNullOrEmpty(supOfProdsByPer.Name))
                                    supOfProdsByPerLst.Add(supOfProdsByPer);

                                if (worksheet.SheetName.Equals(TargetSheets.ANUAL.ToString()))
                                {
                                    actualyear = actualyear + 1;
                                }
                                else
                                { 
                                    quartercount++;
                                    if(quartercount>=5)
                                    {
                                        quartercount = 1;
                                        actualyear = actualyear + 1;
                                    }
                                }
                            }

                            if(!string.IsNullOrEmpty(lastRowName))
                            {
                                if (supOfProdsByPer.Name.Contains(lastRowName) && cellIdx == worksheet.GetRow(rowIdx).LastCellNum-1)
                                    return supOfProdsByPerLst;
                            }


                        }
                    }

            }
            return supOfProdsByPerLst;
        }

        private static double GetFirstYearSuppProd(ISheet worksheet, string dataAtFirstRowFirstCell)
        {
            for (int rowIdx = worksheet.FirstRowNum; rowIdx < worksheet.LastRowNum; rowIdx++)
            {
                
                if (worksheet.GetRow(rowIdx) != null && worksheet.GetRow(rowIdx).FirstCellNum != -1)
                {

                    if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue.Equals(dataAtFirstRowFirstCell))
                        {
                            if(worksheet.SheetName==TargetSheets.ANUAL.ToString())
                            {
                                double year = worksheet.GetRow(rowIdx - 1).GetCell(worksheet.GetRow(rowIdx - 1).FirstCellNum+1).NumericCellValue;
                                if(year>1000.0)
                                return worksheet.GetRow(rowIdx - 1).GetCell(worksheet.GetRow(rowIdx - 1).FirstCellNum+1).NumericCellValue;
                                else
                                { 
                                    year = worksheet.GetRow(rowIdx - 2).GetCell(worksheet.GetRow(rowIdx - 2).FirstCellNum+1).NumericCellValue;
                                    if (year > 1000.0)
                                        return year;
                                    else
                                        return worksheet.GetRow(rowIdx - 2).GetCell(worksheet.GetRow(rowIdx - 2).FirstCellNum+2).NumericCellValue;
                                }
                            }
                            else
                            {
                                try{
                                    return worksheet.GetRow(rowIdx - 3).GetCell(worksheet.GetRow(rowIdx - 3).FirstCellNum + 1).NumericCellValue;
                                }
                                catch { 
                                    return worksheet.GetRow(rowIdx - 2).GetCell(worksheet.GetRow(rowIdx - 2).FirstCellNum + 2).NumericCellValue;
                                }
                            }
                        }
                    }
                }
            }
            throw new Exception("Provided first row data not found");
        }

        private static string GetFirstYearSuppPetProd(ISheet worksheet, string dataAtFirstRowFirstCell)
        {
            for (int rowIdx = worksheet.FirstRowNum; rowIdx < worksheet.LastRowNum; rowIdx++)
            {

                if (worksheet.GetRow(rowIdx) != null && worksheet.GetRow(rowIdx).FirstCellNum != -1)
                {

                    if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue.Equals(dataAtFirstRowFirstCell))
                        {
                            if (worksheet.SheetName == TargetSheets.ANUAL.ToString())
                            {
                                return worksheet.GetRow(rowIdx - 2).GetCell(worksheet.GetRow(rowIdx - 2).FirstCellNum + 1).NumericCellValue.ToString();
                            }
                            else
                            {
                                return worksheet.GetRow(rowIdx - 2).GetCell(worksheet.GetRow(rowIdx - 2).FirstCellNum + 1).StringCellValue;
                            }
                                
                        }
                    }
                }
            }
            throw new Exception("Provided first row data not found");
        }


        private static int GetFirstRowSuppProd(ISheet worksheet, string dataAtFirstRowFirstCell)
        {
            for (int rowIdx = worksheet.FirstRowNum; rowIdx < worksheet.LastRowNum; rowIdx++)
            {

                if (worksheet.GetRow(rowIdx) != null && worksheet.GetRow(rowIdx).FirstCellNum != -1)
                {

                    if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        Console.Write(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue + "    ");
                        if(worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).StringCellValue.Equals(dataAtFirstRowFirstCell))
                        { 
                            return rowIdx;
                        }
                    }
                    if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).CellType == NPOI.SS.UserModel.CellType.Numeric)
                    {
                        if (worksheet.GetRow(rowIdx).GetCell(worksheet.GetRow(rowIdx).FirstCellNum).NumericCellValue== double.Parse(dataAtFirstRowFirstCell))
                        {
                            return rowIdx;
                        }
                    }
                }
            }
            throw new Exception("Provided first row data not found");
        }

    }
}
