using Data_Collector.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Data_Collector
{
    static public class WorkWithDB
    {
        static public List<HorizontalPositionFormat> GetHorizontalPositionPoints (int cycleNuber, List<string> pointsName)
        {
            List<HorizontalPositionFormat> result = new List<HorizontalPositionFormat>();
            using (Context context = new Context())
            {

                foreach (var name in pointsName)
                {
                    var res1 = context.pointCoordinates.SingleOrDefault(point => point.Name == name & point.CycleNumber == cycleNuber);
                    var res2 = context.pointCoordinates.SingleOrDefault(point => point.Name == name & point.CycleNumber == cycleNuber + 1);
                    var res3 = context.pointCoordinates.SingleOrDefault(point => point.Name == name & point.CycleNumber == cycleNuber + 2);
                    if (res3 != null)
                    {
                        result.Add(new HorizontalPositionFormat
                        {
                            Name = name,
                            FirstDateTime = res1.DateTime,
                            FirstNorth = res1.North,
                            FirstEast = res1.East,
                            FirstDiffNorth = res1.NorthDiff,
                            FirstDiffEast = res1.EastDiff,
                            SecondDateTime = res2.DateTime,
                            SecondNorth = res2.North,
                            SecondEast = res2.East,
                            SecondDiffNorth = res2.NorthDiff,
                            SecondDiffEast = res2.EastDiff,
                            ThridDateTime = res3.DateTime,
                            ThridNorth = res3.North,
                            ThridEast = res3.East,
                            ThridDiffNorth = res3.NorthDiff,
                            ThridDiffEast = res3.EastDiff,
                        });
                    }
                    else if (res2 != null)
                    {
                        result.Add(new HorizontalPositionFormat
                        {
                            Name = name,
                            FirstDateTime = res1.DateTime,
                            FirstNorth = res1.North,
                            FirstEast = res1.East,
                            FirstDiffNorth = res1.NorthDiff,
                            FirstDiffEast = res1.EastDiff,
                            SecondDateTime = res2.DateTime,
                            SecondNorth = res2.North,
                            SecondEast = res2.East,
                            SecondDiffNorth = res2.NorthDiff,
                            SecondDiffEast = res2.EastDiff,
                        });
                    }
                    else
                    {
                        result.Add(new HorizontalPositionFormat
                        {
                            Name = name,
                            FirstDateTime = res1.DateTime,
                            FirstNorth = res1.North,
                            FirstEast = res1.East,
                            FirstDiffNorth = res1.NorthDiff,
                            FirstDiffEast = res1.EastDiff,
                        });
                    }
                }
            }
            return result;
        }
        static public List<PositionCoordinate> GetVerticalPositionPoints (int cycleNuber, List<string> pointsName)
        {
            var result = new List<PositionCoordinate>();

            using(Context context = new Context())
            {
                foreach (var name in pointsName)
                {
                    var res1 = context.pointCoordinates.SingleOrDefault(point => point.Name == name & point.CycleNumber == cycleNuber);
                    if(res1 != null)
                    {
                        result.Add(new PositionCoordinate() 
                        {
                            Name = name,
                            DataTime = res1.DateTime,
                            North = res1.North,
                            East = res1.East,
                            NorthDiff  = res1.NorthDiff,
                            EastDiff = res1.EastDiff,
                            Number = res1.CycleNumber
                        });
                    }
                }
            }

            return result;
        }
        static public List<string> GetAllPointsName(string prefix)
        {
            using (Context context = new Context())
            {
                List<string> result = new List<string>();
                foreach (var item in context.pointCoordinates.Select(point => point.Name).Where(name => name.StartsWith(prefix) & name.Contains(prefix)).GroupBy(name => name).Select(name => name.Key).ToList())
                {
                    result.Add(item);
                }
                result.Sort(new NaturalSort());
                return result;
            }
        }

    }
}
