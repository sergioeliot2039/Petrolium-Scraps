using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using ScraperModel.Models;

namespace ScraperModel
{
    public class ScraperDbContext : DbContext
    {
        public void SaveData<T>(List<T> data, DateTime lastModifiedTimestamp) where T : ScrapeTable
        {
            if (data != null && data.Count > 0)
            {
                data.ForEach(d => d.LastModifiedTimestamp = lastModifiedTimestamp);
                data.ForEach(d => d.RunTimestamp = DateTime.Now);
                var dbSet = this.Set<T>();

                // Only uncomment these if you have created a local database;
                //dbSet.AddOrUpdate(data.ToArray());
                //this.SaveChanges();
            }
        }

        public DbSet<GovUkSupplyUseCrudeGasFeedstocks> GovUkSupplyUseCrudeGasFeedstocks { get; set; }
        public DbSet<GovUSAShortTermEnergyOutlook> GovUSAShortTermEnergyOutlook { get; set; }

        public DbSet<GovUSAEnergySummary> GovUSAEnergySummary { get; set; }


    }
}
