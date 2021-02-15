using Data_Collector.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excell = Microsoft.Office.Interop.Excel;

namespace Data_Collector
{
    public class Excel : Data
    {
        Excell.Application ExApp;
        Excell.Workbook Wb;
        Excell.Worksheet Wsh;
        public Excel(Excell.Application ExApp)
        {
            this.ExApp = ExApp;
        }

        public override List<PointCoordinate> GetData()
        {
            List<PointCoordinate> result = new List<PointCoordinate>();
            try
            {
                var DirectoriesPath = Directory.GetFiles("D:\\Отчетные документы").ToList();
                //DirectoriesPath.Clear();
                //for (int i = 88; i < 100; i++)
                //{
                    
                //    DirectoriesPath.Add($"D:\\Отчетные документы\\Таблица к журналу_{i}-й цикл.xlsx");
                //}

                foreach (var a in DirectoriesPath)
                {
                    var path = a.Substring(a.IndexOf('$') + 1);
                    Wb = ExApp.Workbooks.Open(path);
                    Wsh = Wb.Sheets["Таблица для отчета"];
                    int lastRow = Wsh.Cells.SpecialCells(Excell.XlCellType.xlCellTypeLastCell).Row;
                    int zpStartRow = (Wsh.Range[Wsh.Cells[1, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).Find("Зем.полотно").Row;
                    int zpEndRow = (Wsh.Range[Wsh.Cells[zpStartRow + 5, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).End[Excell.XlDirection.xlDown].Row;
                    var bpStartRow = (Wsh.Range[Wsh.Cells[1, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).Find("Балластная призма")?.Row;
                    int? bpEndRow = null, vspEndRow = null;
                    if (bpStartRow != null)
                        bpEndRow = (Wsh.Range[Wsh.Cells[bpStartRow + 5, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).End[Excell.XlDirection.xlDown]?.Row;
                    var vspStartRow = (Wsh.Range[Wsh.Cells[1, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).Find("ВСП")?.Row;
                    if (vspStartRow != null)
                        vspEndRow = (Wsh.Range[Wsh.Cells[vspStartRow + 5, 1], Wsh.Cells[lastRow, 7]] as Excell.Range).End[Excell.XlDirection.xlDown]?.Row;

                    result.AddRange(AllCoordinate(zpStartRow + 5, zpEndRow));

                    if (bpStartRow != null)
                    {
                        result.AddRange(HeightCoordinate(bpStartRow + 5, bpEndRow));
                    }
                    if (vspStartRow != 0)
                    {
                        result.AddRange(AllCoordinate(vspStartRow + 5, vspEndRow));
                    }
                    Wb?.Close(false);
                    Wb = null;
                    Wsh = null;
                    this.SavePoints(result);
                    result.Clear();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                Wb?.Close(false);
                Wb = null;
                Wsh = null;
            }

            return result;
        }
        private List<PointCoordinate> AllCoordinate(int? startRow, int? endRow)
        {
            List<PointCoordinate> result = new List<PointCoordinate>();
            if (startRow != null && endRow != null)
            {
                var Values = (Wsh.Range[Wsh.Cells[startRow, 1], Wsh.Cells[endRow, 7]] as Excell.Range).Cast<Excell.Range>().Select(x => x.Value).ToList();
                int count = 0;
                PointCoordinate point = new PointCoordinate();
                foreach (var item in Values)
                { 
                    switch (count)
                    {
                        case 0:
                            {
                                point.North = item;
                                count++;
                            }
                            break;
                        case 1:
                            {
                                point.East = item;
                                count++;
                            }
                            break;
                        case 2:
                            {
                                point.Height = item;
                                count++;
                            }
                            break;
                        case 3:
                            {
                                point.NorthDiff = item;
                                count++;
                            }
                            break;
                        case 4:
                            {
                                point.EastDiff = item;
                                count++;
                            }
                            break;
                        case 5:
                            {
                                point.HeightDiff = item;
                                count++;
                            }
                            break;
                        case 6:
                            {
                                count = 0;
                                result.Add(new PointCoordinate() 
                                {
                                    North = point.North, 
                                    East = point.East, 
                                    Height = point.Height,
                                    NorthDiff = point.NorthDiff, 
                                    EastDiff = point.EastDiff,
                                    HeightDiff = point.HeightDiff,
                                    Name = item,
                                    DateTime = Convert.ToDateTime((Wsh.Cells[startRow - 3, 1] as Excell.Range).Value),
                                    CycleNumber = (int)(Wsh.Cells[startRow - 4, 1] as Excell.Range).Value
                                });
                            }
                            break;
                    }
                }
            }
            return result;
        }
        private List<PointCoordinate> HeightCoordinate(int? startRow, int? endRow)
        {
            List<PointCoordinate> result = new List<PointCoordinate>();
            if (startRow != null && endRow != null)
            {
                var Values = (Wsh.Range[Wsh.Cells[startRow, 1], Wsh.Cells[endRow, 3]] as Excell.Range).Cast<Excell.Range>().Select(x => x.Value).ToList();
                int count = 0;
                PointCoordinate point = new PointCoordinate();
                foreach (var item in Values)
                {
                    switch (count)
                    {
                        case 0:
                            {
                                point.Height = item;
                                count++;
                            }
                            break;
                        case 1:
                            {
                                point.HeightDiff = item;
                                count++;
                            }
                            break;
                        case 2:
                            {
                                count = 0;
                                result.Add(new PointCoordinate()
                                {
                                    Height = point.Height,
                                    HeightDiff = point.HeightDiff,
                                    Name = item,
                                    DateTime = Convert.ToDateTime((Wsh.Cells[startRow - 3, 1] as Excell.Range).Value),
                                    CycleNumber = (int)(Wsh.Cells[startRow - 4, 1] as Excell.Range).Value
                            });
                            }
                            break;
                    }
                }
            }
            return result;
        }
    }
}
