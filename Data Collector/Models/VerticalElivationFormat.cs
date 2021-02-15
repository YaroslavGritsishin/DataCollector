using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector.Models
{
  public  class VerticalElivationFormat
    {
        public string Name { get; set; }
        public DateTime FirstDateTime { get; set; }
        public double FirstHeight { get; set; }
        public double FirstDiffHeight { get; set; }
        public DateTime SecondDateTime { get; set; }
        public double SecondHeight { get; set; }
        public double SecondDiffHeight { get; set; }
        public DateTime ThridDateTime { get; set; }
        public double ThirdHeghit { get; set; }
        public double ThirdDiffHeight { get; set; }

    }
}
