using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Threading.Tasks;

namespace ScraperModel.Models
{
    public class GovUSAEnergySummary: ScrapeTable
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


        [Required]
        public double Quantity { get; set; }

        [Required]
        public string Source { get; set; }

    }
}
