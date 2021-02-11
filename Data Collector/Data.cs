using Data_Collector.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Collector
{
    public abstract class Data
    {
        public abstract List<PointCoordinate> GetData();
        public void SavePoints(List<PointCoordinate> points)
        {
            using (Context context = new Context())
            {
                context.pointCoordinates.AddRange(points);
                context.SaveChanges();
            }
        }
        public void SaveBasePoint(List<BasePointCoordinate> basePoints)
        {
            using (Context context = new Context())
            {
                context.basePointCoordinates.AddRange(basePoints);
                context.SaveChanges();
            }
        }

    }
}
