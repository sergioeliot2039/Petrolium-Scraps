using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ScraperModel.Models
{
    public class GovUSAShortTermEnergyOutlook : ScrapeTable
    {
        [Key]
        [Column(Order = 0)]
        public DateTime StartDate { get; set; }

        [Key]
        [Column(Order = 1)]
        public DateTime EndDate { get; set; }

        [Key]
        [Column(Order = 2)]
        public string Name { get; set; }

        [Key]
        [Column(Order = 3)]
        public string Code { get; set; }

        [Required]
        public double Quantity { get; set; }

        [Required]
        public string Source { get; set; }


    }
}
