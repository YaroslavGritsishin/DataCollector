using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector.Models
{
  public  class HorizontalElivationFormat
    {
        public string Name { get; set; }
        public DateTime FirstDateTime { get; set; }
        public double FirstHeight { get; set; }
        public double FirstDiffHeight { get; set; }

        public DateTime SecondDateTime { get; set; }
        public double SecondHeight { get; set; }
        public double SecondDiffHeight { get; set; }

        public DateTime ThridDateTime { get; set; }
        public double ThridHeight { get; set; }
        public double ThridDiffHeight { get; set; }

        public DateTime FourDateTime { get; set; }
        public double FourHeight { get; set; }
        public double FourDiffHeight { get; set; }
        
        public DateTime FiveDateTime { get; set; }
        public double FiveHeight { get; set; }
        public double FiveDiffHeight { get; set; }
       
        public DateTime SixDateTime { get; set; }
        public double SixHeight { get; set; }
        public double SixDiffHeight { get; set; }


    }
}
