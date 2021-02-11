using System;

namespace Data_Collector.Models
{
   public class PointCoordinate
    {
        public int Id { get; set; }
        public DateTime DateTime { get; set; }
        public int CycleNumber { get; set; }
        public string Name { get; set; }
        public double North { get; set; }
        public double East { get; set; }
        public double Height { get; set; }
        public double NorthDiff { get; set; }
        public double EastDiff { get; set; }
        public double HeightDiff{ get; set; }

    }
}
