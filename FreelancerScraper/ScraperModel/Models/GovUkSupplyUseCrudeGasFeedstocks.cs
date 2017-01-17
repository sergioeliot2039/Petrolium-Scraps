using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ScraperModel.Models
{
    public class GovUkSupplyUseCrudeGasFeedstocks : ScrapeTable
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
