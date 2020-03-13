using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace RichStringSample
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream("data/sample.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var book = WorkbookFactory.Create(fs);
                var sheet = book.GetSheetAt(0);

                CellReference cellRef = new CellReference("A1");
                var cell = sheet.GetRow(cellRef.Row).GetCell(cellRef.Col);

                var xssfStr = (XSSFRichTextString)cell.RichStringCellValue;

                for (int runNo = 0; runNo < xssfStr.NumFormattingRuns; runNo++)
                {
                    var xssfFont = (XSSFFont)xssfStr.GetFontOfFormattingRun(runNo);

                    var idx = xssfStr.GetIndexOfFormattingRun(runNo);
                    var len = xssfStr.GetLengthOfFormattingRun(runNo);

                    Console.WriteLine(
                        $"RunNo:{string.Format("{0:D2}", runNo)} 「{xssfStr.String.Substring(idx, len)}」" +
                        Environment.NewLine + " =>" +
                        $"Font:{xssfFont.FontName}, " +
                        $"Color:{xssfFont.GetXSSFColor()?.GetARGBHex()}, " +
                        $"Size:{xssfFont.FontHeightInPoints}, " +
                        $"Bold:{string.Format("{0, 5}", xssfFont.IsBold)}, " +
                        $"Italic:{string.Format("{0, 5}", xssfFont.IsItalic)}, " +
                        $"Underline:{string.Format("{0, 6}", xssfFont.Underline)}, " +
                        $"Strikeout:{string.Format("{0, 5}", xssfFont.IsStrikeout)}, " +
                        $"TypeOffset:{string.Format("{0, 5}", xssfFont.TypeOffset)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (fs != null) fs.Dispose();
            }
        }
    }
}