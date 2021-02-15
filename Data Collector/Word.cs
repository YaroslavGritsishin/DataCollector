using System;
using System.Collections.Generic;
using Data_Collector.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Data_Collector
{
    static public class Word
    { 
        static public int ShetsCount { get; set; }
        static public Table CreatePositionTableWithHeight(List<PositionCoordinate> points)
        {
            Table reportTable = new Table();
            TableProperties tblProp = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
               ),
                    new TableWidth() { Width = WidthForWord(17.18), Type = TableWidthUnitValues.Dxa },
                    new TableJustification() { Val = TableRowAlignmentValues.Center },
                    new TableLayout() { Type = TableLayoutValues.Fixed },
                    new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true }
                    );
            reportTable.Append(tblProp);
            TableGrid grid = new TableGrid(
                new GridColumn() { Width = WidthForWord(0.9) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.25) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.25) },
                new GridColumn() { Width = WidthForWord(1.69) },
                new GridColumn() { Width = WidthForWord(1.69) },
                new GridColumn() { Width = WidthForWord(2.1) }
                );
            reportTable.Append(grid);
            //Добавление Шапки Таблицы 
            foreach (TableRow row in PositionHeader(points[0].Number, points[0].DataTime, new DateTime(2018, 6, 1)))
            {
                reportTable.Append(row);
            }
            //Наполнение таблицы
            int count = 1;
            foreach (var point in points)
            {
                reportTable.Append(CreateRow(point, count,0.5));
                count++;
            }

            return reportTable;
        }
        static public Table CreateVerticalPositionTable(List<PositionCoordinate> points)
        {
            Table reportTable = new Table();
            TableProperties tblProp = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
               ),
                    new TableWidth() { Width = WidthForWord(17.18), Type = TableWidthUnitValues.Dxa },
                    new TableJustification() { Val = TableRowAlignmentValues.Center },
                    new TableLayout() { Type = TableLayoutValues.Fixed },
                    new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true }
                    );
            reportTable.Append(tblProp);
            TableGrid grid = new TableGrid(
                new GridColumn() { Width = WidthForWord(0.9) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.25) },
                new GridColumn() { Width = WidthForWord(2.1) },
                new GridColumn() { Width = WidthForWord(2.25) },
                new GridColumn() { Width = WidthForWord(1.69) },
                new GridColumn() { Width = WidthForWord(1.69) },
                new GridColumn() { Width = WidthForWord(2.1) }
                );
            reportTable.Append(grid);
            //Добавление Шапки Таблицы 
            foreach (TableRow row in PositionHeader(points[0].Number, points[0].DataTime, new DateTime(2018, 6, 1)))
            {
                reportTable.Append(row);
            }
            //Наполнение таблицы
            int count = 1;
            foreach (var point in points)
            {
                reportTable.Append(CreateRow(point, count, 0.35));
                count++;
            }

            return reportTable;
        }
        static public Table CreteHorizontalPositionTable(List<HorizontalPositionFormat> points, int number, double rowHeight)
        {
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;
            Table reportTable = new Table(new TableProperties().Width(28).AddBorders(BorderValues.Single, 1));
            TableGrid grid = new TableGrid(
                new GridColumn() { Width = WidthForWord(numericColumnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) });
            reportTable.Append(grid);
            foreach (var row in HorizontalPositionHeader(number, points[0].FirstDateTime, points[0].SecondDateTime, points[0].ThridDateTime))
            {
                reportTable.Append(row);
            }
            int count = 1;
            foreach (var point in points)
            {
                reportTable.Append(CreateRow(point, count, rowHeight));
                count++;
            }


            return reportTable;
        }
        static public Table CreateVerticalElivationTable(List<VerticalElivationFormat> points, int number)
        {
            Table reportTable = new Table(new TableProperties().Width(15.9).AddBorders(BorderValues.Single, 1));
            TableGrid grid = new TableGrid(
                new GridColumn() { Width = WidthForWord(0.96) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) },
                new GridColumn() { Width = WidthForWord(1.66) });
            reportTable.Append(grid);
            foreach (var row in VerticalElivationHeader(number, points[0].FirstDateTime, points[0].SecondDateTime, points[0].ThridDateTime))
            {
                reportTable.Append(row);
            }
            int count = 1;
            foreach (var point in points)
            {
                reportTable.Append(CreateRow(point, count));
                count++;
            }


            return reportTable;
        }
        static public Table CreteHorizontalElevationTable(List<HorizontalElivationFormat> points, int number, double rowHeight)
        {
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;
            Table reportTable = new Table(new TableProperties().Width(26.23).AddBorders(BorderValues.Single, 1));
            TableGrid grid = new TableGrid(
                new GridColumn() { Width = WidthForWord(numericColumnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) },
                new GridColumn() { Width = WidthForWord(columnWidth) });
            reportTable.Append(grid);
            foreach (var row in HorizontalElivationHeader(number, points[0].FirstDateTime, points[0].SecondDateTime, points[0].ThridDateTime, points[0].FourDateTime, points[0].FiveDateTime, points[0].SixDateTime))
            {
                reportTable.Append(row);
            }
            int count = 1;
            foreach (var point in points)
            {
                reportTable.Append(CreateRow(point, count, rowHeight));
                count++;
            }


            return reportTable;
        }
        
        static private TableRow CreateRow(PositionCoordinate point, int count, double rowHeght)
        {
            string Size = "15";

            TableRow row = new TableRow(new TableRowProperties(
                new TableJustification() { Val = TableRowAlignmentValues.Center }
                //new TableRowHeight() { Val = Convert.ToUInt32("175"), HeightType = HeightRuleValues.Exact }
            ).Heihgt(rowHeght));



            TableCell countCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(0.9) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));

            countCell.Append(new Paragraph(AlignText(new ParagraphProperties()),
                new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text($"{count}"))));
            row.Append(countCell);


            TableCell nameCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            nameCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text($"{point.Name}"))));
            row.Append(nameCell);


            TableCell baseNorthCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));

            string calcBaseNorth = string.Empty;
            if (point.NorthDiff < 0) calcBaseNorth = (point.North - point.NorthDiff / 1000).ToString("0.0000");
            else calcBaseNorth = (point.North - point.NorthDiff / 1000).ToString("0.0000");
            baseNorthCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(calcBaseNorth))));
            row.Append(baseNorthCell);

            string calcBaseEast = string.Empty;
            if (point.EastDiff < 0) calcBaseEast = (point.East - point.EastDiff / 1000).ToString("0.0000");
            else calcBaseEast = (point.East - point.EastDiff / 1000).ToString("0.0000");

            TableCell baseEastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.25) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            baseEastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(calcBaseEast))));
            row.Append(baseEastCell);


            TableCell northCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            northCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(point.North.ToString("0.0000")))));
            row.Append(northCell);

            TableCell eastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.25) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            eastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(point.East.ToString("0.0000")))));
            row.Append(eastCell);

            TableCell diffNorthCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(1.69) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            diffNorthCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(point.NorthDiff.ToString("0.0")))));
            row.Append(diffNorthCell);

            TableCell diffEastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(1.69) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            diffEastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text(point.EastDiff.ToString("0.0")))));
            row.Append(diffEastCell);

            TableCell lastNameCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ));
            lastNameCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(new RunProperties(
                    new FontSize() { Val = Size }),
                    new Text($"{point.Name}"))));
            row.Append(lastNameCell);

            return row;
        }
        static private TableRow CreateRow(VerticalElivationFormat point, int count)
        {
            TableRow row = new TableRow(new TableRowProperties().Heihgt(0.31).TextAlign(TableRowAlignmentValues.Center));

            TableCell countCell = new TableCell(new TableCellProperties().Width(0.6).VerticalAligmet(TableVerticalAlignmentValues.Center));
            countCell.AddValue(count.ToString(), fontSize: 7.5);
            row.Append(countCell);

            TableCell nameCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue(point.Name, fontSize: 7.5);
            row.Append(nameCell);

            TableCell baseHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseHeightCell.AddValue((point.FirstHeight - point.FirstDiffHeight / 1000).ToString("0.0000"), fontSize: 7.5);
            row.Append(baseHeightCell);

            TableCell firstHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstHeightCell.AddValue(point.FirstHeight.ToString("0.0000"), fontSize: 7.5);
            row.Append(firstHeightCell);

            TableCell firstDiffHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffHeightCell.AddValue(point.FirstDiffHeight.ToString("0.0"), fontSize: 7.5);
            row.Append(firstDiffHeightCell);

            TableCell secondHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondHeightCell.AddValue(point.SecondHeight.ToString("0.0000"), fontSize: 7.5);
            row.Append(secondHeightCell);

            TableCell secondDiffHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffHeightCell.AddValue(point.SecondDiffHeight.ToString("0.0"), fontSize: 7.5);
            row.Append(secondDiffHeightCell);

            TableCell thirdHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thirdHeightCell.AddValue(point.ThirdHeghit.ToString("0.0000"), fontSize: 7.5);
            row.Append(thirdHeightCell);

            TableCell thirdDiffHeightCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thirdDiffHeightCell.AddValue(point.ThirdDiffHeight.ToString("0.0"), fontSize: 7.5);
            row.Append(thirdDiffHeightCell);

            TableCell nameEndCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue(point.Name, fontSize: 7.5);
            row.Append(nameEndCell);

            return row;
        }
        static private TableRow CreateRow(HorizontalPositionFormat point, int count,double rowHeght)
        {
            double fontSize = 7.5;
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;

            TableRow row = new TableRow(new TableRowProperties().Heihgt(rowHeght).TextAlign(TableRowAlignmentValues.Center));

            TableCell countCell = new TableCell(new TableCellProperties().Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            countCell.AddValue(count.ToString(), fontSize: fontSize);
            row.Append(countCell);

            TableCell nameCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue(point.Name, fontSize: fontSize);
            row.Append(nameCell);

            TableCell basePositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            basePositionXCell.AddValue((point.FirstNorth - point.FirstDiffNorth / 1000).ToString("0.0000"), fontSize: fontSize);
            row.Append(basePositionXCell);

            TableCell basePositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            basePositionYCell.AddValue((point.FirstEast - point.FirstDiffEast / 1000).ToString("0.0000"), fontSize: fontSize);
            row.Append(basePositionYCell);

            // Первый блок
            TableCell firstPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstPositionXCell.AddValue(point.FirstNorth.ToString("0.0000"), fontSize: fontSize);
            row.Append(firstPositionXCell);

            TableCell firstPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstPositionYCell.AddValue(point.FirstEast.ToString("0.0000"), fontSize: fontSize);
            row.Append(firstPositionYCell);

            TableCell firstDiffPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffPositionXCell.AddValue(point.FirstDiffNorth.ToString("0.0"), fontSize: fontSize);
            row.Append(firstDiffPositionXCell);

            TableCell firstDiffPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffPositionYCell.AddValue(point.FirstDiffEast.ToString("0.0"), fontSize: fontSize);
            row.Append(firstDiffPositionYCell);

            // Второй блок
            TableCell secondPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondPositionXCell.AddValue(point.SecondNorth.ToString("0.0000"), fontSize: fontSize);
            row.Append(secondPositionXCell);

            TableCell secondPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondPositionYCell.AddValue(point.SecondEast.ToString("0.0000"), fontSize: fontSize);
            row.Append(secondPositionYCell);

            TableCell secondDiffPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffPositionXCell.AddValue(point.SecondDiffNorth.ToString("0.0"), fontSize: fontSize);
            row.Append(secondDiffPositionXCell);

            TableCell secondDiffPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffPositionYCell.AddValue(point.SecondDiffEast.ToString("0.0"), fontSize: fontSize);
            row.Append(secondDiffPositionYCell);

            //Третий блок
            TableCell thridPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridPositionXCell.AddValue(point.ThridNorth.ToString("0.0000"), fontSize: fontSize);
            row.Append(thridPositionXCell);

            TableCell thridPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridPositionYCell.AddValue(point.ThridEast.ToString("0.0000"), fontSize: fontSize);
            row.Append(thridPositionYCell);

            TableCell thridDiffPositionXCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridDiffPositionXCell.AddValue(point.ThridDiffNorth.ToString("0.0"), fontSize: fontSize);
            row.Append(thridDiffPositionXCell);

            TableCell thridDiffPositionYCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridDiffPositionYCell.AddValue(point.ThridDiffEast.ToString("0.0"), fontSize: fontSize);
            row.Append(thridDiffPositionYCell);

            TableCell nameEndCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue(point.Name, fontSize: fontSize);
            row.Append(nameEndCell);

            return row;
        }
        static private TableRow CreateRow(HorizontalElivationFormat point, int count, double rowHeght)
        {
            double fontSize = 7.5;
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;

            TableRow row = new TableRow(new TableRowProperties().Heihgt(rowHeght).TextAlign(TableRowAlignmentValues.Center));

            TableCell countCell = new TableCell(new TableCellProperties().Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            countCell.AddValue(count.ToString(), fontSize: fontSize);
            row.Append(countCell);

            TableCell nameCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue(point.Name, fontSize: fontSize);
            row.Append(nameCell);

            TableCell baseElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseElevationCell.AddValue((point.FirstHeight - point.FirstDiffHeight / 1000).ToString("0.0000"), fontSize: fontSize);
            row.Append(baseElevationCell);

            // Первый блок
            TableCell firstElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstElevationCell.AddValue(point.FirstHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(firstElevationCell);

            TableCell firstDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffElevationCell.AddValue(point.FirstDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(firstDiffElevationCell);

            // Второй блок
            TableCell secondElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondElevationCell.AddValue(point.SecondHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(secondElevationCell);

            TableCell secondDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffElevationCell.AddValue(point.SecondDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(secondDiffElevationCell);


            //Третий блок
            TableCell thridElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridElevationCell.AddValue(point.ThridHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(thridElevationCell);

            TableCell thridDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridDiffElevationCell.AddValue(point.ThridDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(thridDiffElevationCell);

            //Четвертый блок
            TableCell fourElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourElevationCell.AddValue(point.FourHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(fourElevationCell);

            TableCell fourDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourDiffElevationCell.AddValue(point.FourDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(fourDiffElevationCell);


            //Пятый блок
            TableCell fiveElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveElevationCell.AddValue(point.FiveHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(fiveElevationCell);

            TableCell fiveDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveDiffElevationCell.AddValue(point.FiveDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(fiveDiffElevationCell);


            //Шестой блок
            TableCell sixElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixElevationCell.AddValue(point.SixHeight.ToString("0.0000"), fontSize: fontSize);
            row.Append(sixElevationCell);

            TableCell sixDiffElevationCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixDiffElevationCell.AddValue(point.SixDiffHeight.ToString("0.0"), fontSize: fontSize);
            row.Append(sixDiffElevationCell);


            TableCell nameEndCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue(point.Name, fontSize: fontSize);
            row.Append(nameEndCell);

            return row;
        }
        static public void AddTableInBookMark(string pathTemplate, Table table, string bookMarkName)
        {

            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathTemplate, true))
            {
                // Append the table to the document.
                //doc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run()));
                //doc.MainDocumentPart.Document.Body.Append(table);
                IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

                foreach (BookmarkStart bookmarkStart in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
                {
                    bookmarkMap[bookmarkStart.Name] = bookmarkStart;
                }

                foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
                {
                    if (bookmarkStart.Name == bookMarkName)
                    {
                        
                        var parent = bookmarkStart.Parent;
                        Paragraph paragraph = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                        parent.InsertBeforeSelf(table);
                        table.InsertAfterSelf(paragraph);
                    }

                    /*Run bookmarkText = bookmarkStart.NextSibling<Run>();
                    if (bookmarkText != null)
                    {
                        bookmarkText.GetFirstChild<Text>().Text = "Збс";
                    }*/
                }
                doc.Save();
                ShetsCount = Convert.ToInt32(doc.ExtendedFilePropertiesPart.Properties.Pages.Text);
            }

        }
        static public void AddTableInBookMark(string pathTemplate, Table table, string bookMarkName, int nexString = 1)
        {

            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathTemplate, true))
            {
                // Append the table to the document.
                //doc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run()));
                //doc.MainDocumentPart.Document.Body.Append(table);
                IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

                foreach (BookmarkStart bookmarkStart in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
                {
                    bookmarkMap[bookmarkStart.Name] = bookmarkStart;
                }

                foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
                {
                    if (bookmarkStart.Name == bookMarkName)
                    {

                        var parent = bookmarkStart.Parent;
                        parent.InsertBeforeSelf(table);
                        if (nexString == 2)
                        {
                            Paragraph paragraph = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                            table.InsertAfterSelf(paragraph);
                        }
                        else
                        {
                            Paragraph paragraph = new Paragraph(new Run());
                            table.InsertAfterSelf(paragraph);
                        }
                          
                    }

                    /*Run bookmarkText = bookmarkStart.NextSibling<Run>();
                    if (bookmarkText != null)
                    {
                        bookmarkText.GetFirstChild<Text>().Text = "Збс";
                    }*/
                }
                doc.Save();
                ShetsCount = Convert.ToInt32(doc.ExtendedFilePropertiesPart.Properties.Pages.Text);
            }

        }

        static private string WidthForWord(double value)
        {
            return (value * 567).ToString();
        }
        static private ParagraphProperties AlignText(ParagraphProperties paragraph)
        {
            paragraph.Append(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
            );
            return paragraph;
        }
        static private List<TableRow> PositionHeader(int number, DateTime dateTime, DateTime baseDateTime)
        {
            string Size = "18";
            List<TableRow> result = new List<TableRow>();
            #region Первый ряд
            TableRow FirstRow = new TableRow(new TableRowProperties(
                new TableJustification() { Val = TableRowAlignmentValues.Center },
                new TableRowHeight() { Val = Convert.ToUInt32("283"), HeightType = HeightRuleValues.Exact }
            ));

            TableCell numberSumbolCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(0.9) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            numberSumbolCell.Append(new Paragraph(AlignText(new ParagraphProperties()),
                new Run(new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()
                ), new Text("№"))));
            FirstRow.Append(numberSumbolCell);

            TableCell nameCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            nameCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text("Назв. деф. марки"))));
            FirstRow.Append(nameCell);

            TableCell baseCycleNumberCell = new TableCell(new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new GridSpan() { Val = 2 }
                ));
            baseCycleNumberCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"0-й цикл"))));
            FirstRow.Append(baseCycleNumberCell);

            TableCell cycleNumberCell = new TableCell(new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new GridSpan() { Val = 4 }
                ));
            cycleNumberCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"{number}-й цикл"))));
            FirstRow.Append(cycleNumberCell);

            TableCell lastNameCell = new TableCell(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
               new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
               new VerticalMerge() { Val = MergedCellValues.Restart }
               ));
            lastNameCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()),
                    new Text("Назв. деф. марки"))));
            FirstRow.Append(lastNameCell);

            result.Add(FirstRow);
            #endregion

            #region Второй ряд
            TableRow SecondRow = new TableRow(new TableRowProperties(
                new TableJustification() { Val = TableRowAlignmentValues.Center },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new TableRowHeight() { Val = Convert.ToUInt32("283"), HeightType = HeightRuleValues.Exact }
            ));

            TableCell numberSumbolCell_2 = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(0.9) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                ));
            numberSumbolCell_2.Append(new Paragraph(new Run(new RunProperties(
                    new FontSize() { Val = Size }))));
            SecondRow.Append(numberSumbolCell_2);

            TableCell nameCell_2 = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                ));
            nameCell_2.Append(new Paragraph(new Run(new RunProperties(
                    new FontSize() { Val = Size }))));
            SecondRow.Append(nameCell_2);

            TableCell baseDateTimeCell = new TableCell(new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new GridSpan() { Val = 2 }
                ));
            baseDateTimeCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"{baseDateTime.ToString("dd.MM.yyyy")}"))));
            SecondRow.Append(baseDateTimeCell);

            TableCell dateTimeCell = new TableCell(new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new GridSpan() { Val = 4 }
                ));
            dateTimeCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"{dateTime.ToString("dd.MM.yyyy HH:mm")}"))));
            SecondRow.Append(dateTimeCell);

            TableCell lastNameCell_2 = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                ));
            lastNameCell_2.Append(new Paragraph(new Run()));
            SecondRow.Append(lastNameCell_2);

            result.Add(SecondRow);
            #endregion

            #region Третий ряд
            TableRow ThirdRow = new TableRow(new TableRowProperties(
               new TableJustification() { Val = TableRowAlignmentValues.Center },
               new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
               new TableRowHeight() { Val = Convert.ToUInt32("283"), HeightType = HeightRuleValues.Exact }
           ));

            TableCell numberSumbolCell_3 = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(0.9) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                ));
            numberSumbolCell_3.Append(new Paragraph(new Run(new RunProperties(
                    new FontSize() { Val = Size }))));
            ThirdRow.Append(numberSumbolCell_3);

            TableCell nameCell_3 = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                ));
            nameCell_3.Append(new Paragraph(new Run(new RunProperties(
                    new FontSize() { Val = Size }))));
            ThirdRow.Append(nameCell_3);

            TableCell baseNorthCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            baseNorthCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"X0, м"))));
            ThirdRow.Append(baseNorthCell);

            TableCell baseEastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.25) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            baseEastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"Y0, м"))));
            ThirdRow.Append(baseEastCell);


            TableCell northCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            northCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"X{number}, м"))));
            ThirdRow.Append(northCell);

            TableCell eastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.25) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            eastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"Y{number}, м"))));
            ThirdRow.Append(eastCell);


            TableCell diffNorthCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(1.69) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            diffNorthCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"{char.ConvertFromUtf32(916)}X{number}, мм"))));
            ThirdRow.Append(diffNorthCell);

            TableCell diffEastCell = new TableCell(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(1.69) },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge() { Val = MergedCellValues.Restart }
                ));
            diffEastCell.Append(new Paragraph(AlignText(new ParagraphProperties()), new Run(
                    new RunProperties(
                    new FontSize() { Val = Size },
                    new Bold()), new Text($"{char.ConvertFromUtf32(916)}Y{number}, мм"))));
            ThirdRow.Append(diffEastCell);

            TableCell lastNameCell_3 = new TableCell(new TableCellProperties(
                 new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = WidthForWord(2.1) },
                 new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                 new VerticalMerge()
                ));
            lastNameCell_3.Append(new Paragraph(new Run(new RunProperties(
                    new FontSize() { Val = Size }))));
            ThirdRow.Append(lastNameCell_3);

            result.Add(ThirdRow);
            #endregion

            #region Создание пусто ряда
            TableRow FourRow = new TableRow(new TableRowProperties(
               new TableJustification() { Val = TableRowAlignmentValues.Center },
               new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
               new TableRowHeight() { Val = Convert.ToUInt32("283"), HeightType = HeightRuleValues.Exact }
           ));
            for (int i = 0; i < 9; i++)
            {
                TableCell Cell = new TableCell(new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new VerticalMerge()
                )); ;
                Cell.Append(new Paragraph(new Run()));
                FourRow.Append(Cell);
            }
            result.Add(FourRow);
            #endregion

            return result;
        }

        static private List<TableRow> VerticalElivationHeader(int number, DateTime firstDateTime, DateTime secondDateTime, DateTime thridDateTime)
        {
            List<TableRow> result = new List<TableRow>();
            double mainColumnWidth = 0.96;
            double columnWidth = 1.66;

            #region Первый ряд
            TableRow firstRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(mainColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell.AddValue("№", bold: true, fontSize: 8);
            firstRow.Append(numberSumbolCell);

            TableCell nameCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameCell);

            TableCell baseCycleNumberCell = new TableCell(new TableCellProperties().Width(1.7).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseCycleNumberCell.AddValue($"0-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(baseCycleNumberCell);

            TableCell firstCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleNumberCell.AddValue($"{number}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(firstCycleNumberCell);

            TableCell secondCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleNumberCell.AddValue($"{number + 1}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(secondCycleNumberCell);

            TableCell thridCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleNumberCell.AddValue($"{number + 2}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(thridCycleNumberCell);

            TableCell nameEndCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameEndCell);
            result.Add(firstRow);
            #endregion

            #region Второй ряд
            TableRow secondRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(mainColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(numberSumbolCell_2);

            TableCell nameCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameCell_2);

            TableCell baseHeightCell_2 = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseHeightCell_2.AddValue("01.06.2018", bold: true, fontSize: 8);
            secondRow.Append(baseHeightCell_2);

            TableCell firstCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleDateTimeCell.AddValue(firstDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(firstCycleDateTimeCell);

            TableCell secondCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleDateTimeCell.AddValue(secondDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(secondCycleDateTimeCell);

            TableCell thridCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleDateTimeCell.AddValue(thridDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(thridCycleDateTimeCell);

            TableCell nameEndCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameEndCell_2);
            result.Add(secondRow);
            #endregion

            #region  Третий ряд
            TableRow thridRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(mainColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(numberSumbolCell_3);

            TableCell nameCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameCell_3);

            TableCell baseHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseHeightCell.AddValue("H0, м", bold: true, fontSize: 8);
            thridRow.Append(baseHeightCell);

            TableCell firstHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstHeightCell.AddValue($"H{number}", bold: true, fontSize: 8);
            thridRow.Append(firstHeightCell);

            TableCell firstDiffHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffHeightCell.AddValue($"{char.ConvertFromUtf32(916)}H{number} - H0", bold: true, fontSize: 8);
            thridRow.Append(firstDiffHeightCell);

            TableCell secondHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondHeightCell.AddValue($"H{number + 1}", bold: true, fontSize: 8);
            thridRow.Append(secondHeightCell);

            TableCell secondDiffHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffHeightCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 1} - H0", bold: true, fontSize: 8);
            thridRow.Append(secondDiffHeightCell);

            TableCell thridHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridHeightCell.AddValue($"H{number + 2}", bold: true, fontSize: 8);
            thridRow.Append(thridHeightCell);

            TableCell thridDiffHeightCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridDiffHeightCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 2} - H0", bold: true, fontSize: 8);
            thridRow.Append(thridDiffHeightCell);

            TableCell nameEndCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameEndCell_3);
            result.Add(thridRow);
            #endregion

            #region Создание пусто ряда
            TableRow FourRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));
            for (int i = 0; i < 11; i++)
            {
                TableCell Cell = new TableCell(new TableCellProperties(new VerticalMerge()).VerticalAligmet(TableVerticalAlignmentValues.Center));
                Cell.Append(new Paragraph(new Run()));
                FourRow.Append(Cell);
            }
            result.Add(FourRow);
            #endregion

            return result;
        }
        static private List<TableRow> HorizontalPositionHeader(int number, DateTime firstDateTime, DateTime secondDateTime, DateTime thridDateTime)
        {
            List<TableRow> result = new List<TableRow>();
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;

            #region Первый ряд
            TableRow firstRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell.AddValue("№", bold: true, fontSize: 8);
            firstRow.Append(numberSumbolCell);

            TableCell nameCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameCell);

            TableCell baseCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseCycleNumberCell.AddValue($"0-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(baseCycleNumberCell);

            TableCell firstCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleNumberCell.AddValue($"{number}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(firstCycleNumberCell);

            TableCell secondCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleNumberCell.AddValue($"{number + 1}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(secondCycleNumberCell);

            TableCell thridCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleNumberCell.AddValue($"{number + 2}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(thridCycleNumberCell);

            TableCell nameEndCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameEndCell);
            result.Add(firstRow);
            #endregion

            #region Второй ряд
            TableRow secondRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(numberSumbolCell_2);

            TableCell nameCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameCell_2);

            TableCell baseHeightCell_2 = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseHeightCell_2.AddValue("01.06.2018", bold: true, fontSize: 8);
            secondRow.Append(baseHeightCell_2);

            TableCell firstCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleDateTimeCell.AddValue(firstDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(firstCycleDateTimeCell);

            TableCell secondCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleDateTimeCell.AddValue(secondDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(secondCycleDateTimeCell);

            TableCell thridCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 4 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleDateTimeCell.AddValue(thridDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(thridCycleDateTimeCell);

            TableCell nameEndCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameEndCell_2);
            result.Add(secondRow);
            #endregion

            #region  Третий ряд
            TableRow thridRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(numberSumbolCell_3);

            TableCell nameCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameCell_3);

            TableCell basePositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            basePositionXCell.AddValue("X0, м", bold: true, fontSize: 8);
            thridRow.Append(basePositionXCell);

            TableCell basePositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            basePositionYCell.AddValue("Y0, м", bold: true, fontSize: 8);
            thridRow.Append(basePositionYCell);

            // Первый блок
            TableCell firstPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstPositionXCell.AddValue($"X{number}", bold: true, fontSize: 8);
            thridRow.Append(firstPositionXCell);

            TableCell firstPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstPositionYCell.AddValue($"Y{number}", bold: true, fontSize: 8);
            thridRow.Append(firstPositionYCell);

            TableCell firstDiffPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffPositionXCell.AddValue($"{char.ConvertFromUtf32(916)}X{number} - X0", bold: true, fontSize: 8);
            thridRow.Append(firstDiffPositionXCell);

            TableCell firstDiffPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffPositionYCell.AddValue($"{char.ConvertFromUtf32(916)}Y{number} - Y0", bold: true, fontSize: 8);
            thridRow.Append(firstDiffPositionYCell);

            //2-й блок
            TableCell SecondPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            SecondPositionXCell.AddValue($"X{number + 1}", bold: true, fontSize: 8);
            thridRow.Append(SecondPositionXCell);

            TableCell SecondPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            SecondPositionYCell.AddValue($"Y{number + 1}", bold: true, fontSize: 8);
            thridRow.Append(SecondPositionYCell);

            TableCell SecondDiffPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            SecondDiffPositionXCell.AddValue($"{char.ConvertFromUtf32(916)}X{number + 1} - X0", bold: true, fontSize: 8);
            thridRow.Append(SecondDiffPositionXCell);

            TableCell SecondDiffPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            SecondDiffPositionYCell.AddValue($"{char.ConvertFromUtf32(916)}Y{number + 1} - Y0", bold: true, fontSize: 8);
            thridRow.Append(SecondDiffPositionYCell);

            //3-й блок
            TableCell ThridPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            ThridPositionXCell.AddValue($"X{number + 2}", bold: true, fontSize: 8);
            thridRow.Append(ThridPositionXCell);

            TableCell ThridPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            ThridPositionYCell.AddValue($"Y{number + 2}", bold: true, fontSize: 8);
            thridRow.Append(ThridPositionYCell);

            TableCell ThridDiffPositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            ThridDiffPositionXCell.AddValue($"{char.ConvertFromUtf32(916)}X{number + 2} - X0", bold: true, fontSize: 8);
            thridRow.Append(ThridDiffPositionXCell);

            TableCell ThridDiffPositionYCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            ThridDiffPositionYCell.AddValue($"{char.ConvertFromUtf32(916)}Y{number + 2} - Y0", bold: true, fontSize: 8);
            thridRow.Append(ThridDiffPositionYCell);

            //Название точки
            TableCell nameEndCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameEndCell_3);
            result.Add(thridRow);
            #endregion

            #region Создание пусто ряда
            TableRow FourRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));
            for (int i = 0; i < 18; i++)
            {
                TableCell Cell = new TableCell(new TableCellProperties(new VerticalMerge()).VerticalAligmet(TableVerticalAlignmentValues.Center));
                Cell.Append(new Paragraph(new Run()));
                FourRow.Append(Cell);
            }
            result.Add(FourRow);
            #endregion

            return result;
        }
        static private List<TableRow> HorizontalElivationHeader(int number, DateTime firstDateTime, DateTime secondDateTime, DateTime thridDateTime, DateTime fourDateTime, DateTime fiveDateTime, DateTime sixDateTime)
        {
            List<TableRow> result = new List<TableRow>();
            double numericColumnWidth = 0.9;
            double columnWidth = 1.77;

            #region Первый ряд
            TableRow firstRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell.AddValue("№", bold: true, fontSize: 8);
            firstRow.Append(numberSumbolCell);

            TableCell nameCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameCell);

            TableCell baseCycleNumberCell = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseCycleNumberCell.AddValue($"0-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(baseCycleNumberCell);

            TableCell firstCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleNumberCell.AddValue($"{number}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(firstCycleNumberCell);

            TableCell secondCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleNumberCell.AddValue($"{number + 1}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(secondCycleNumberCell);

            TableCell thridCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleNumberCell.AddValue($"{number + 2}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(thridCycleNumberCell);

            TableCell fourCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourCycleNumberCell.AddValue($"{number + 3}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(fourCycleNumberCell);

            TableCell fiveCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveCycleNumberCell.AddValue($"{number + 4}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(fiveCycleNumberCell);

            TableCell sixCycleNumberCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixCycleNumberCell.AddValue($"{number + 5}-й Цикл", bold: true, fontSize: 8);
            firstRow.Append(sixCycleNumberCell);

            TableCell nameEndCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell.AddValue("Назв. деф. марки", bold: true, fontSize: 8);
            firstRow.Append(nameEndCell);
            result.Add(firstRow);
            #endregion

            #region Второй ряд
            TableRow secondRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(numberSumbolCell_2);

            TableCell nameCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameCell_2);

            TableCell baseHeightCell_2 = new TableCell(new TableCellProperties().Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            baseHeightCell_2.AddValue("01.06.2018", bold: true, fontSize: 8);
            secondRow.Append(baseHeightCell_2);

            TableCell firstCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstCycleDateTimeCell.AddValue(firstDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(firstCycleDateTimeCell);

            TableCell secondCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondCycleDateTimeCell.AddValue(secondDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(secondCycleDateTimeCell);

            TableCell thridCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridCycleDateTimeCell.AddValue(thridDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(thridCycleDateTimeCell);

            TableCell fourCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourCycleDateTimeCell.AddValue(fourDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(fourCycleDateTimeCell);

            TableCell fiveCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveCycleDateTimeCell.AddValue(fiveDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(fiveCycleDateTimeCell);

            TableCell sixCycleDateTimeCell = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixCycleDateTimeCell.AddValue(sixDateTime.ToString("dd-MM-yyyy HH:mm"), bold: true, fontSize: 8);
            secondRow.Append(sixCycleDateTimeCell);

            TableCell nameEndCell_2 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_2.Append(new Paragraph(new Run()));
            secondRow.Append(nameEndCell_2);
            result.Add(secondRow);
            #endregion

            #region  Третий ряд
            TableRow thridRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));

            TableCell numberSumbolCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(numericColumnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            numberSumbolCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(numberSumbolCell_3);

            TableCell nameCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameCell_3);

            TableCell basePositionXCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            basePositionXCell.AddValue("H0, м", bold: true, fontSize: 8);
            thridRow.Append(basePositionXCell);

            // Первый блок
            TableCell firstElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstElevationCell.AddValue($"H{number}", bold: true, fontSize: 8);
            thridRow.Append(firstElevationCell);

            TableCell firstDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            firstDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number} - H0", bold: true, fontSize: 8);
            thridRow.Append(firstDiffElevationCell);

            //Второй блок
            TableCell secondElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondElevationCell.AddValue($"H{number + 1}", bold: true, fontSize: 8);
            thridRow.Append(secondElevationCell);

            TableCell secondDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            secondDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 1} - H0", bold: true, fontSize: 8);
            thridRow.Append(secondDiffElevationCell);

            //Третий блок
            TableCell thridElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridElevationCell.AddValue($"H{number + 2}", bold: true, fontSize: 8);
            thridRow.Append(thridElevationCell);

            TableCell thridDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            thridDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 2} - H0", bold: true, fontSize: 8);
            thridRow.Append(thridDiffElevationCell);

            //Четвертый блок
            TableCell fourElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourElevationCell.AddValue($"H{number + 3}", bold: true, fontSize: 8);
            thridRow.Append(fourElevationCell);

            TableCell fourDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fourDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 3} - H0", bold: true, fontSize: 8);
            thridRow.Append(fourDiffElevationCell);

            //Пятый блок
            TableCell fiveElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveElevationCell.AddValue($"H{number + 4}", bold: true, fontSize: 8);
            thridRow.Append(fiveElevationCell);

            TableCell fiveDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            fiveDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 4} - H0", bold: true, fontSize: 8);
            thridRow.Append(fiveDiffElevationCell);

            //Шестой блок
            TableCell sixElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixElevationCell.AddValue($"H{number + 5}", bold: true, fontSize: 8);
            thridRow.Append(sixElevationCell);

            TableCell sixDiffElevationCell = new TableCell(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart }).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            sixDiffElevationCell.AddValue($"{char.ConvertFromUtf32(916)}H{number + 5} - H0", bold: true, fontSize: 8);
            thridRow.Append(sixDiffElevationCell);

            //Название точки
            TableCell nameEndCell_3 = new TableCell(new TableCellProperties(new VerticalMerge()).Width(columnWidth).VerticalAligmet(TableVerticalAlignmentValues.Center));
            nameEndCell_3.Append(new Paragraph(new Run()));
            thridRow.Append(nameEndCell_3);
            result.Add(thridRow);
            #endregion

            #region Создание пусто ряда
            TableRow FourRow = new TableRow(new TableRowProperties().Heihgt(0.5).TextAlign(TableRowAlignmentValues.Center));
            for (int i = 0; i < 17; i++)
            {
                TableCell Cell = new TableCell(new TableCellProperties(new VerticalMerge()).VerticalAligmet(TableVerticalAlignmentValues.Center));
                Cell.Append(new Paragraph(new Run()));
                FourRow.Append(Cell);
            }
            result.Add(FourRow);
            #endregion

            return result;
        }

    }
}
