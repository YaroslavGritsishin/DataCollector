using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector.Models
{
   public class PositionCoordinate
    {
        public int Id { get; set; }
        public int Number { get; set; }
        public string Name { get; set; }
        public double North { get; set; }
        public double East { get; set; }
        public double Height { get; set; }
        public double NorthDiff { get; set; }
        public double EastDiff { get; set; }
        public double HeightDiff { get; set; }
        public DateTime DataTime { get; set; }

    }
}
