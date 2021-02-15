using Data_Collector.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;


namespace Data_Collector
{
    [SuppressUnmanagedCodeSecurity]
   public class NaturalSort: IComparer, IComparer<ElivationCoordinate>, IComparer<string>, IComparer<PositionCoordinate>, IComparer<VerticalElivationFormat>, IComparer<HorizontalPositionFormat>, IComparer<HorizontalElivationFormat>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string psz1, string psz2);
        public int Compare(ElivationCoordinate x, ElivationCoordinate y)
        {
            return StrCmpLogicalW(x.Name, y.Name);
        }

        public int Compare(string x, string y)
        {
            return StrCmpLogicalW(x, y);
        }

        public int Compare(PositionCoordinate x, PositionCoordinate y)
        {
            return StrCmpLogicalW(x.Name, y.Name);
        }

        public int Compare(VerticalElivationFormat x, VerticalElivationFormat y)
        {
            return StrCmpLogicalW(x.Name, y.Name);
        }

        public int Compare(HorizontalPositionFormat x, HorizontalPositionFormat y)
        {
            return StrCmpLogicalW(x.Name, y.Name);
        }

        public int Compare(HorizontalElivationFormat x, HorizontalElivationFormat y)
        {
            return StrCmpLogicalW(x.Name, y.Name);
        }

        public int Compare(object x, object y)
        {
            return StrCmpLogicalW(x.ToString(), y.ToString());
        }
    }
}
