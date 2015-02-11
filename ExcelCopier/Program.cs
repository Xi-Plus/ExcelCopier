using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Text.RegularExpressions;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;


namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream file = new FileStream(@"in.xlsx", FileMode.Open);

            var workbook = new XSSFWorkbook(file);
            var sheet = workbook.GetSheetAt(0);

            var workbookout = new XSSFWorkbook();
            var sheetout = workbookout.CreateSheet("out");

            sheetout.SetMargin(MarginType.HeaderMargin, sheet.GetMargin(MarginType.HeaderMargin));
            sheetout.SetMargin(MarginType.FooterMargin, sheet.GetMargin(MarginType.FooterMargin));
            sheetout.SetMargin(MarginType.RightMargin, sheet.GetMargin(MarginType.RightMargin));
            sheetout.SetMargin(MarginType.TopMargin, sheet.GetMargin(MarginType.TopMargin));
            sheetout.SetMargin(MarginType.LeftMargin, sheet.GetMargin(MarginType.LeftMargin));
            sheetout.SetMargin(MarginType.BottomMargin, sheet.GetMargin(MarginType.BottomMargin));

            char[] delimiterChars = { ',', '\t' };
            int count = 0;
            int rowcount = sheet.LastRowNum+1;
            foreach (string line in File.ReadLines("in.csv", Encoding.Default))
            {
                string[] text = line.Split(delimiterChars);
                int textreplacecount = 0;
                for (int rowNo = 0; rowNo <= rowcount; rowNo++)
                {
                    var row = sheet.GetRow(rowNo);
                    if (row == null)
                        continue;
                    sheetout.CreateRow(rowNo + count * rowcount);
                    sheetout.CreateRow(rowNo + count * rowcount).HeightInPoints = sheet.GetRow(rowNo).HeightInPoints;
                    
                    for (int cellNo = 0; cellNo <= row.LastCellNum; cellNo++)
                    {
                        var cell = row.GetCell(cellNo);
                        if (cell == null) // null is when the cell is empty
                            continue;
                        if (cell.IsMergedCell)
                        {
                            //Console.WriteLine(string.Format("Cell row: {0} column: {1} IsMergedCell ", cell.Row.RowNum, cell.ColumnIndex));
                        }
                        string s = cell.StringCellValue;
                        while (true)
                        {
                            string type=null;
                            int index=2147483647;
                            if (s.IndexOf("{}") != -1 && s.IndexOf("{}") < index)
                            {
                                type = "{}";
                                index = s.IndexOf("{}");
                            }
                            if (s.IndexOf("{s}") != -1 && s.IndexOf("{s}") < index)
                            {
                                type = "{s}";
                                index = s.IndexOf("{s}");
                                text[textreplacecount] = numtochinese(int.Parse(text[textreplacecount]));
                            }
                            if (s.IndexOf("{n2}") != -1 && s.IndexOf("{n2}") < index)
                            {
                                type = "{n2}";
                                index = s.IndexOf("{n2}");
                                text[textreplacecount] = numfillzero(text[textreplacecount]);
                            }
                            if (s.IndexOf("{class}") != -1 && s.IndexOf("{class}") < index)
                            {
                                type = "{class}";
                                index = s.IndexOf("{class}");
                                text[textreplacecount] = classtostring(text[textreplacecount]);
                            }
                            if (type != null)
                            {
                                Regex rgx = new Regex(type);
                                s = rgx.Replace(s, text[textreplacecount], 1);
                            }
                            else break;
                            //Console.WriteLine(s);
                            textreplacecount++;
                        }
                        sheetout.GetRow(rowNo + count * rowcount).CreateCell(cellNo).SetCellValue(s);
                        sheetout.GetRow(rowNo + count * rowcount).GetCell(cellNo).CellStyle.CloneStyleFrom(cell.CellStyle);

                        ICellStyle cellStyle = null;
                        cellStyle = workbookout.CreateCellStyle();
                        IFont font = null;
                        font = workbookout.CreateFont();
                        font.FontHeightInPoints = cell.CellStyle.GetFont(workbook).FontHeightInPoints;
                        font.FontName = cell.CellStyle.GetFont(workbook).FontName;
                        font.Boldweight = cell.CellStyle.GetFont(workbook).Boldweight;
                        cellStyle.Alignment = cell.CellStyle.Alignment;
                        cellStyle.VerticalAlignment = cell.CellStyle.VerticalAlignment;
                        cellStyle.SetFont(font);
                        //sheetout.GetRow(rowNo).GetCell(cellNo).CellStyle.CloneStyleFrom(cellStyle);
                        sheetout.GetRow(rowNo + count * rowcount).GetCell(cellNo).CellStyle = cellStyle;

                        //Console.WriteLine(string.Format("Cell row: {0} column: {1} has value: {2}", cell.Row.RowNum, cell.ColumnIndex, cell.StringCellValue));
                        //Console.WriteLine(string.Format("Cell row: {0} column: {1} has style: {2}", cell.Row.RowNum, cell.ColumnIndex, cell.CellStyle.GetFont(workbook).FontHeightInPoints));

                    }
                }
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    CellRangeAddress mergedRegion = sheet.GetMergedRegion(i);
                    var r1 = mergedRegion.FirstRow + count * rowcount;
                    var r2 = mergedRegion.LastRow + count * rowcount;
                    var c1 = mergedRegion.FirstColumn;
                    var c2 = mergedRegion.LastColumn;
                    sheetout.AddMergedRegion(new CellRangeAddress(r1, r2, c1, c2));
                }
                count++;
                sheetout.SetRowBreak(count * rowcount - 1);
            }
            for (int cellNo = 0; cellNo <= sheet.GetRow(0).LastCellNum; cellNo++)
            {
                sheetout.SetColumnWidth(cellNo, (int)(sheet.GetColumnWidth(cellNo)*1.09));
            }
            

            FileStream fileout = new FileStream(@"out.xlsx", FileMode.Create);//產生檔案
            workbookout.Write(fileout);

            fileout.Close();
            Console.ReadKey();
        }

        private static string numfillzero(string p)
        {
            if (p.Length < 2) p = "0" + p;
            return p;
        }

        private static string classtostring(string p)
        {
            int n = int.Parse(p);
            string s = "";
            s += numtochinese(n / 100) + "年";
            s += numtochinese(n % 100) + "班";
            return s;
        }

        private static string numtochinese(int n)
        {
            string[] num = { "", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };
            string s = "";
            s += num[n / 10 * 10];
            s += num[n % 10];
            return s;
        }
    }
}
