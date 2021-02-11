using System.Data.Entity;

namespace Data_Collector.Models
{
    public class Context: DbContext
    {
        public DbSet<PointCoordinate> pointCoordinates { get; set; }
        public DbSet<BasePointCoordinate> basePointCoordinates { get; set; }
        public Context():base("Connect")
        {
            Database.CreateIfNotExists();
        }
    }
}
