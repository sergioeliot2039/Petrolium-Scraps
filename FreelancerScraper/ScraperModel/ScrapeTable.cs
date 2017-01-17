using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ScraperModel
{
    public class ScrapeTable
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        
        [Required]
        public DateTime LastModifiedTimestamp { get; set; }

        [Required]
        public DateTime RunTimestamp { get; set; }

    }
}
