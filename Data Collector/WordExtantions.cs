using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Data_Collector
{
    public static class WordExtantions
    {
        public enum TextPosition
        {
            Left,
            Center,
            Right
        }
        public static ParagraphProperties TextAlign(this ParagraphProperties paragraphProperties, TextPosition textPosition)
        {
            if (textPosition == TextPosition.Center)
            {
                paragraphProperties.Append(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
            }
            if (textPosition == TextPosition.Left)
            {
                paragraphProperties.Append(
                new Justification() { Val = JustificationValues.Start },
                new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
            }
            if (textPosition == TextPosition.Right)
            {
                paragraphProperties.Append(
                new Justification() { Val = JustificationValues.End },
                new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
            }
            return paragraphProperties;
        }
        public static RunProperties FontFamily(this RunProperties runProperties, string fontName)
        {
            runProperties.Append(new RunFonts() { Ascii = fontName });
            return runProperties;
        }
        public static RunProperties FontSize(this RunProperties runProperties, double value)
        {
            string calcValue = Math.Round(value * 2).ToString();
            runProperties.Append(new FontSize() { Val = calcValue });
            return runProperties;
        }
        public static RunProperties Bold(this RunProperties runProperties)
        {
            runProperties.Append(new Bold());
            return runProperties;
        }
        public static RunProperties AddText(this RunProperties runProperties, string value)
        {
            runProperties.Append(new Text(value));
            return runProperties;
        }

        public static TableProperties AddBorders(this TableProperties tableProperties, BorderValues borderValues, UInt32 size)
        {
            tableProperties.Append(new TableBorders(
                     new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size },
                     new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size },
                     new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size },
                     new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size },
                     new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size },
                     new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size }));

            return tableProperties;
        }
        public static TableProperties Width (this TableProperties tableProperties, double width, TableRowAlignmentValues tableRowAlignmentValues = TableRowAlignmentValues.Center, TableLayoutValues tableLayoutValues = TableLayoutValues.Fixed)
        {
            string calcValue = (width * 567).ToString();
            tableProperties.Append(new TableWidth() { Width = calcValue, Type = TableWidthUnitValues.Dxa },
                    new TableJustification() { Val = TableRowAlignmentValues.Center },
                    new TableLayout() { Type = TableLayoutValues.Fixed },
                    new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true });
            return tableProperties;
        }

        public static TableRowProperties Heihgt(this TableRowProperties tableRowProperties, double value)
        {
            var calcValue = Convert.ToUInt32(value * 567);
            tableRowProperties.Append(new TableRowHeight() { Val = calcValue, HeightType = HeightRuleValues.Exact });
            return tableRowProperties;
        }
        public static TableRowProperties TextAlign(this TableRowProperties tableRowProperties, TableRowAlignmentValues tableRowAlignmentValues)
        {
            tableRowProperties.Append(new TableJustification() { Val = tableRowAlignmentValues });
            return tableRowProperties;
        }


        public static TableCell AddValue(this TableCell tableCell, string value, string fontFamily = "Times New Roman", double fontSize = 12, bool bold = false)
        {
            if (bold)
            {
                tableCell.Append(new Paragraph(new ParagraphProperties().TextAlign(TextPosition.Center),
                    new Run(new RunProperties().FontFamily(fontFamily).FontSize(fontSize).Bold().AddText(value))));
            }
            else
            {
                tableCell.Append(new Paragraph(new ParagraphProperties().TextAlign(TextPosition.Center),
                    new Run(new RunProperties().FontFamily(fontFamily).FontSize(fontSize).AddText(value))));
            }


            return tableCell;
        }
        public static TableCellProperties Width(this TableCellProperties tableCellProperties, double value)
        {
            string calValue = (value * 567).ToString();
            tableCellProperties.Append(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = calValue });
            return tableCellProperties;
        }
        public static TableCellProperties VerticalAligmet(this TableCellProperties tableCellProperties, TableVerticalAlignmentValues tableVerticalAlignmentValues)
        {
            tableCellProperties.Append(new TableCellVerticalAlignment() { Val = tableVerticalAlignmentValues });
            return tableCellProperties;
        }

        public static IEnumerable<(T item, int index)> WithIndex<T>(this IEnumerable<T> source)
        {
            return source.Select((item, index) => (item, index));
        }
    }
}
