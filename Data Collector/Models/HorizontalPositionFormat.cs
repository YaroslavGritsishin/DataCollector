using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector.Models
{
  public  class HorizontalPositionFormat
    {
        public string Name { get; set; }
        public DateTime FirstDateTime { get; set; }
        public double FirstNorth { get; set; }
        public double FirstDiffNorth { get; set; }
        public double FirstEast { get; set; }
        public double FirstDiffEast { get; set; }
        public DateTime SecondDateTime { get; set; }
        public double SecondNorth { get; set; }
        public double SecondDiffNorth { get; set; }
        public double SecondEast { get; set; }
        public double SecondDiffEast { get; set; }
        public DateTime ThridDateTime { get; set; }
        public double ThridNorth { get; set; }
        public double ThridDiffNorth { get; set; }
        public double ThridEast { get; set; }
        public double ThridDiffEast { get; set; }
    }
}
