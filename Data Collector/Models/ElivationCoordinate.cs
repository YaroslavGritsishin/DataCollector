using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector.Models
{
   public class ElivationCoordinate
    {
        public int Id { get; set; }
        public int Number { get; set; }
        public string Name { get; set; }
        public double Height { get; set; }
        public double HeightDiff { get; set; }
        public DateTime DataTime { get; set; }

    }
}
