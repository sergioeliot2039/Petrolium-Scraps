﻿using System;
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
    public class GovUSAElectricityTrade : ExcelXlsxScrapeData
    {
        public const string URL = "https://www.eia.gov/outlooks/aeo/excel/aeotab_10.xlsx";

        protected override void ScrapeXlsx(ScraperDbContext dbContext, DateTime lastModified, XLWorkbook xlsx)
        {

            #region ref2017.1208a
            var petroleumSheet = xlsx.Worksheet("ref2017.1208a");
            var petroleumData = new List<GovUSAShortTermEnergyOutlook>();
            int firstRowWithData = 18;
            var cellShift = 3;
            var rowCellHeaders = 13;
            int lastCellWithData = 39;
            int lastRow = 36;
            string category = "";

            for (int rowIdx = firstRowWithData; rowIdx <= lastRow; rowIdx++)
            {
                var petroleumRow = petroleumSheet.Row(rowIdx);
                var previusRow = petroleumSheet.Row(rowIdx - 1);
                int year = 0;
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
                    if (rowIdx >= 18 && rowIdx <= 20)
                        prefixName = "Interregional Electricity Trade Gross Domestic Sales ";
                    if (rowIdx >= 23 && rowIdx <= 25)
                        prefixName = "Interregional Electricity Trade Gross Domestic Sales (million 2016 dollars) ";
                    if (rowIdx >= 29 && rowIdx <= 31)
                        prefixName = "International Electricity Trade Imports from Canada and Mexico ";
                    if (rowIdx >= 34 && rowIdx <= 36)
                        prefixName = "International Electricity Trade Exports to Canada and Mexico ";

                    petroleumDatum.Code = prefixName + petroleumRow.Cell(1).Value.ToString();                    
                    petroleumDatum.Name = prefixName + petroleumRow.Cell(2).Value.ToString();                    
                    double quantity;
                    bool parseQtyorrect = false;
                    parseQtyorrect = double.TryParse(petroleumRow.Cell(cellIdx).Value.ToString(), out quantity);
                    petroleumDatum.Quantity = parseQtyorrect ? quantity : 0.0;

                    DateTime startDt = new DateTime();
                    DateTime endDt = new DateTime();                    
                    year = int.Parse(petroleumSheet.Row(rowCellHeaders).Cell(cellIdx).Value.ToString());
                    GetStartEndDate(year, "Year", out startDt, out endDt);
                    petroleumDatum.StartDate = startDt;
                    petroleumDatum.EndDate = endDt;
                    //if (rowIdx < 35) continue;
                    petroleumData.Add(petroleumDatum);
                    //Console.WriteLine("Year: {0} Month: {1} ", year, petroleumDatum.StartDate.Month);                    
                    year++;                    

                    //if (!shown) { Console.WriteLine(petroleumDatum.Name); shown = true; }
                }
                //if(rowIdx!=17)break;
            }
            //Console.ReadKey();    
            
            Console.WriteLine("Code \t Name \t Quantity \t Year");
            foreach (GovUSAShortTermEnergyOutlook s in petroleumData)
            {
                Console.WriteLine("{0} \t {1} \t {2}\t {3} ",
                    s.Code, s.Name, s.Quantity, s.StartDate.Year + " " + s.StartDate.Month);
            }
            Console.ReadKey();

            #endregion  ref2017.1208a

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
            else if (month.Contains("Year"))
            {
                StartDt = new DateTime(int.Parse(year.ToString()), 1, 1);
                EndDt = new DateTime(int.Parse(year.ToString()), 12, 31);
            }

            StartDate = StartDt;
            EndDate = EndDt;
        }

    }
}