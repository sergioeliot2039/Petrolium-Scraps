using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScrappingDataFromWebPageInCS
{


    public sealed class TargetSheets
    {

        private readonly String name;
        private readonly int value;

        public static readonly TargetSheets ANUAL = new TargetSheets(1, "Annual");
        public static readonly TargetSheets QUARTER = new TargetSheets(2, "Quarter");
        public static readonly TargetSheets MONTH = new TargetSheets(3, "Month");
        public static readonly TargetSheets MONTH1 = new TargetSheets(4, "Month-1");
        public static readonly TargetSheets MONTH2 = new TargetSheets(5, "Month-2");

        private TargetSheets(int value, String name)
        {
            this.name = name;
            this.value = value;
        }

        public override String ToString()
        {
            return name;
        }
    }

    public class SupplyOfProductsByPeriod
    {

        private string month;

        public string Month
        {
            get { return month; }
            set { month = value;
            
                if(month.Contains("January"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString().ToString()), 1, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 1, 31);
                }
                else if (month.Contains("February"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 2, 1);
                    this.EndDate = DateTime.IsLeapYear(int.Parse(year.ToString())) ? new DateTime(int.Parse(year.ToString()), 2, 29) : new DateTime(int.Parse(year.ToString()), 2, 28);
                }
                else if (month.Contains("March"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 3, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 3, 31);
                }
                else if (month.Contains("April"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 4, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 4, 30);
                }
                else if (month.Contains("May"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 5, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 5, 31);
                }
                else if (month.Contains("June"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 6, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 6, 30);
                }
                else if (month.Contains("July"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 7, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 7, 31);                
                }
                else if (month.Contains("August"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 8, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 8, 31);      
                }
                else if (month.Contains("September"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 9, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 9, 30);      
                }
                else if (month.Contains("October"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 10, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 10, 31);  
                }
                else if (month.Contains("November"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 11, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 11, 30); 
                }
                else if (month.Contains("December"))
                {
                    this.StartDate = new DateTime(int.Parse(year.ToString()), 12, 1);
                    this.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);  
                }
            
            }
        }

        public string Source;

        public DateTime StartDate;

        public DateTime EndDate;

        private double year;

        public double Year
        {
            get { return year; }
            set { 
                year = value;
                this.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                this.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
            }
        }
        private string quarter;

        public string Quarter
        {
            get { return quarter; }
            set { quarter = value;

                switch(quarter)
                { 
                    case "Q1":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 3, 31);
                        break;
                    case "Q2":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 4, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 6, 30);
                        break;
                    case "Q3":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 7, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 9, 30);
                        break;
                    case "Q4":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 10, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        break;
                }

            }
        }
        private double quantity;

        public double Quantity
        {
            get { return quantity; }
            set { quantity = value; }
        }

        private string name;
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
       
    }

    public class SupplyOfPetProductsByPeriod
    {
        public DateTime StartDate;

        public DateTime EndDate;

        public string Source; 

        private string name;
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        private double year;

        public double Year
        {
            get { return year; }
            set
            {
                year = value;
                this.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                this.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
            }
        }
        private string quarter;

        public string Quarter
        {
            get { return quarter; }
            set
            {
                quarter = value;

                switch (quarter)
                {
                    case "Q1":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 3, 31);
                        break;
                    case "Q2":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 4, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 6, 30);
                        break;
                    case "Q3":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 7, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 9, 30);
                        break;
                    case "Q4":
                        this.StartDate = new DateTime(int.Parse(year.ToString()), 10, 1);
                        this.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
                        break;
                }

            }
        }

        private double quantity;

        public double Quantity
        {
            get { return quantity; }
            set { quantity = value; }
        }

        private double totalPetroleumProducts;

        public double TotalPetroleumProducts
        {
            get { return totalPetroleumProducts; }
            set { totalPetroleumProducts = value; }
        }
        private double motorSpirit;

        public double MotorSpirit
        {
            get { return motorSpirit; }
            set { motorSpirit = value; }
        }
        private double dERV;

        public double DERV
        {
            get { return dERV; }
            set { dERV = value; }
        }
        private double gasOil;

        public double GasOil
        {
            get { return gasOil; }
            set { gasOil = value; }
        }
        private double aviationTurbineFuel;

        public double AviationTurbineFuel
        {
            get { return aviationTurbineFuel; }
            set { aviationTurbineFuel = value; }
        }
        private double fuelOils;

        public double FuelOils
        {
            get { return fuelOils; }
            set { fuelOils = value; }
        }
        private double petroleumGases;

        public double PetroleumGases
        {
            get { return petroleumGases; }
            set { petroleumGases = value; }
        }
        private double burningOil;

        public double BurningOil
        {
            get { return burningOil; }
            set { burningOil = value; }
        }
        private double otherProducts;

        public double OtherProducts
        {
            get { return otherProducts; }
            set { otherProducts = value; }
        }

    }
}
