using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Util;
using NPOI.SS.Formula.Functions;
using System.Text;


namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws = wb.CreateSheet("Class");
            ws.SetMargin(MarginType.HeaderMargin, (double)0.8 / 2.5);
            ws.SetMargin(MarginType.FooterMargin, (double)0.8 / 2.5);

            ws.SetMargin(MarginType.RightMargin, (double)0.8 / 2.5);
            ws.SetMargin(MarginType.TopMargin, (double)1.9 / 2.5);
            ws.SetMargin(MarginType.LeftMargin, (double)1.3 / 2.5);
            ws.SetMargin(MarginType.BottomMargin, (double)1.9 / 2.5);

            int count = 0;

            char[] delimiterChars = { ' ', ',', '\t' };

            const int rows = 15;
            string[] num = { "", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };
            foreach (string line in File.ReadLines("in.csv", Encoding.Default))
            {
                string[] rawtext = line.Split(delimiterChars);
                string[] text = new string[12];
                for (int i = 1; i <= 11; i++)
                {
                    text[i] = rawtext[i - 1];
                }
                int cla = int.Parse(rawtext[0]);
                text[0] = num[cla / 100];
                cla %= 100;
                text[1] = num[cla / 10 * 10] + num[cla % 10];
                text[4] = num[int.Parse(text[4])];
                switch (rawtext[4])
                {
                    case "l":
                        text[5] = rawtext[7] + "公尺";
                        break;
                    case "m":
                        text[5] = rawtext[5] + "分" + rawtext[6] + "秒" + rawtext[7];
                        break;
                    case "s":
                        text[5] = rawtext[6] + "秒" + rawtext[7];
                        break;
                }
                for (int i = 0; i <= 11;i++ ){
                    if(i==6||i==7||i==8)continue;
                    Console.Write(text[i] + ',');
                }
                Console.WriteLine("");
                if (text[2].IndexOf("?") != -1) Console.WriteLine("*編碼錯誤 第" + (4 + count * rows) + "行");

                ws.CreateRow(0 + count * rows);
                ws.GetRow(0 + count * rows).HeightInPoints = 39.75F;

                ws.CreateRow(1 + count * rows);
                ws.GetRow(1 + count * rows).HeightInPoints = 67.50F;

                ICellStyle cellStyle1 = null;
                cellStyle1 = wb.CreateCellStyle();
                IFont font1 = null;
                font1 = wb.CreateFont();
                font1.FontHeightInPoints = 48;
                font1.FontName = "標楷體";
                font1.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle1.Alignment = HorizontalAlignment.Center;
                cellStyle1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle1.SetFont(font1);
                ws.GetRow(1 + count * rows).CreateCell(3).SetCellValue("獎　    狀");
                ws.GetRow(1 + count * rows).GetCell(3).CellStyle = cellStyle1;

                ws.CreateRow(2 + count * rows);
                ws.GetRow(2 + count * rows).HeightInPoints = 72.00F;

                ws.CreateRow(3 + count * rows);
                ws.GetRow(3 + count * rows).HeightInPoints = 46.50F;

                ICellStyle cellStyle2 = null;
                cellStyle2 = wb.CreateCellStyle();
                IFont font2 = null;
                font2 = wb.CreateFont();
                font2.FontHeightInPoints = 22;
                font2.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle2.Alignment = HorizontalAlignment.Right;
                cellStyle2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle2.SetFont(font2);
                ws.GetRow(3 + count * rows).CreateCell(2).SetCellValue("本校 ");
                ws.GetRow(3 + count * rows).GetCell(2).CellStyle = cellStyle2;

                ICellStyle cellStyle3 = null;
                cellStyle3 = wb.CreateCellStyle();
                IFont font3 = null;
                font3 = wb.CreateFont();
                font3.FontHeightInPoints = 22;
                font3.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle3.Alignment = HorizontalAlignment.Center;
                cellStyle3.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle3.SetFont(font3);
                ws.GetRow(3 + count * rows).CreateCell(3).SetCellValue(text[0] + "年" + text[1] + "班 ");
                ws.GetRow(3 + count * rows).GetCell(3).CellStyle = cellStyle3;

                ICellStyle cellStyle4 = null;
                cellStyle4 = wb.CreateCellStyle();
                IFont font4 = null;
                font4 = wb.CreateFont();
                font4.FontHeightInPoints = 22;
                font4.FontName = "標楷體";
                font4.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle4.Alignment = HorizontalAlignment.Center;
                cellStyle4.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle4.SetFont(font4);
                ws.GetRow(3 + count * rows).CreateCell(4).SetCellValue(text[2]);
                ws.GetRow(3 + count * rows).GetCell(4).CellStyle = cellStyle4;

                ICellStyle cellStyle5 = null;
                cellStyle5 = wb.CreateCellStyle();
                IFont font5 = null;
                font5 = wb.CreateFont();
                font5.FontHeightInPoints = 22;
                font5.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle5.Alignment = HorizontalAlignment.Left;
                cellStyle5.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle5.SetFont(font5);
                ws.GetRow(3 + count * rows).CreateCell(5).SetCellValue("同學");
                ws.GetRow(3 + count * rows).GetCell(5).CellStyle = cellStyle5;

                ws.CreateRow(4 + count * rows);
                ws.GetRow(4 + count * rows).HeightInPoints = 46.50F;

                ICellStyle cellStyle6 = null;
                cellStyle6 = wb.CreateCellStyle();
                IFont font6 = null;
                font6 = wb.CreateFont();
                font6.FontHeightInPoints = 22;
                font6.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle6.Alignment = HorizontalAlignment.Center;
                cellStyle6.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle6.SetFont(font6);
                ws.GetRow(4 + count * rows).CreateCell(1).SetCellValue("參加103學年度全校運動會");
                ws.GetRow(4 + count * rows).GetCell(1).CellStyle = cellStyle6;

                ICellStyle cellStyle7 = null;
                cellStyle7 = wb.CreateCellStyle();
                IFont font7 = null;
                font7 = wb.CreateFont();
                font7.FontHeightInPoints = 22;
                font7.FontName = "標楷體";
                font7.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle7.Alignment = HorizontalAlignment.Center;
                cellStyle7.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle7.SetFont(font7);
                ws.GetRow(4 + count * rows).CreateCell(4).SetCellValue(text[3]);
                ws.GetRow(4 + count * rows).GetCell(4).CellStyle = cellStyle7;

                ws.CreateRow(5 + count * rows);
                ws.GetRow(5 + count * rows).HeightInPoints = 46.50F;

                ICellStyle cellStyle8 = null;
                cellStyle8 = wb.CreateCellStyle();
                IFont font8 = null;
                font8 = wb.CreateFont();
                font8.FontHeightInPoints = 22;
                font8.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle8.Alignment = HorizontalAlignment.Left;
                cellStyle8.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle8.SetFont(font8);
                ws.GetRow(5 + count * rows).CreateCell(1).SetCellValue("榮獲");
                ws.GetRow(5 + count * rows).GetCell(1).CellStyle = cellStyle8;

                ICellStyle cellStyle9 = null;
                cellStyle9 = wb.CreateCellStyle();
                IFont font9 = null;
                font9 = wb.CreateFont();
                font9.FontHeightInPoints = 22;
                font9.FontName = "標楷體";
                font9.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle9.Alignment = HorizontalAlignment.Left;
                cellStyle9.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle9.SetFont(font9);
                ws.GetRow(5 + count * rows).CreateCell(2).SetCellValue("第" + text[4] + "名");
                ws.GetRow(5 + count * rows).GetCell(2).CellStyle = cellStyle9;

                ICellStyle cellStyle10 = null;
                cellStyle10 = wb.CreateCellStyle();
                IFont font10 = null;
                font10 = wb.CreateFont();
                font10.FontHeightInPoints = 22;
                font10.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle10.Alignment = HorizontalAlignment.Center;
                cellStyle10.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle10.SetFont(font10);
                ws.GetRow(5 + count * rows).CreateCell(3).SetCellValue("成績");
                ws.GetRow(5 + count * rows).GetCell(3).CellStyle = cellStyle10;

                ICellStyle cellStyle11 = null;
                cellStyle11 = wb.CreateCellStyle();
                IFont font11 = null;
                font11 = wb.CreateFont();
                font11.FontHeightInPoints = 22;
                font11.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle11.Alignment = HorizontalAlignment.Left;
                cellStyle11.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle11.SetFont(font11);
                ws.GetRow(5 + count * rows).CreateCell(4).SetCellValue(text[5]);
                ws.GetRow(5 + count * rows).GetCell(4).CellStyle = cellStyle11;

                ws.CreateRow(6 + count * rows);
                ws.GetRow(6 + count * rows).HeightInPoints = 54.75F;
                ws.CreateRow(7 + count * rows);
                ws.GetRow(7 + count * rows).HeightInPoints = 30.00F;

                ICellStyle cellStyle12 = null;
                cellStyle12 = wb.CreateCellStyle();
                IFont font12 = null;
                font12 = wb.CreateFont();
                font12.FontHeightInPoints = 22;
                font12.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle12.Alignment = HorizontalAlignment.Left;
                cellStyle12.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle12.SetFont(font12);
                ws.GetRow(7 + count * rows).CreateCell(1).SetCellValue("特頒獎狀 以資表揚");
                ws.GetRow(7 + count * rows).GetCell(1).CellStyle = cellStyle12;

                ws.CreateRow(8 + count * rows);
                ws.GetRow(8 + count * rows).HeightInPoints = 30.00F;
                ws.CreateRow(9 + count * rows);
                ws.GetRow(9 + count * rows).HeightInPoints = 30.00F;
                ws.CreateRow(10 + count * rows);
                ws.GetRow(10 + count * rows).HeightInPoints = 87.75F;
                ws.CreateRow(11 + count * rows);
                ws.GetRow(11 + count * rows).HeightInPoints = 60.75F;
                ws.CreateRow(12 + count * rows);
                ws.GetRow(12 + count * rows).HeightInPoints = 69.75F;
                ws.CreateRow(13 + count * rows);
                ws.GetRow(13 + count * rows).HeightInPoints = 30.00F;

                ICellStyle cellStyle13 = null;
                cellStyle13 = wb.CreateCellStyle();
                IFont font13 = null;
                font13 = wb.CreateFont();
                font13.FontHeightInPoints = 22;
                font13.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle13.Alignment = HorizontalAlignment.Left;
                cellStyle13.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle13.SetFont(font13);
                ws.GetRow(13 + count * rows).CreateCell(1).SetCellValue("中 華 民 國");
                ws.GetRow(13 + count * rows).GetCell(1).CellStyle = cellStyle13;

                ICellStyle cellStyle14 = null;
                cellStyle14 = wb.CreateCellStyle();
                IFont font14 = null;
                font14 = wb.CreateFont();
                font14.FontHeightInPoints = 22;
                font14.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle14.Alignment = HorizontalAlignment.Left;
                cellStyle14.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle14.SetFont(font14);
                ws.GetRow(13 + count * rows).CreateCell(3).SetCellValue(text[9] + "年  " + text[10] + "月");
                ws.GetRow(13 + count * rows).GetCell(3).CellStyle = cellStyle14;

                ICellStyle cellStyle15 = null;
                cellStyle15 = wb.CreateCellStyle();
                IFont font15 = null;
                font15 = wb.CreateFont();
                font15.FontHeightInPoints = 22;
                font15.FontName = "標楷體";
                //font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyle15.Alignment = HorizontalAlignment.Left;
                cellStyle15.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyle15.SetFont(font15);
                ws.GetRow(13 + count * rows).CreateCell(4).SetCellValue(text[11] + "日");
                ws.GetRow(13 + count * rows).GetCell(4).CellStyle = cellStyle15;

                ws.CreateRow(14 + count * rows);
                ws.SetRowBreak(ws.GetRow(14 + count * rows).CreateCell(4).RowIndex);

                ws.AddMergedRegion(new CellRangeAddress(4 + count * rows, 4 + count * rows, 1, 3));
                ws.AddMergedRegion(new CellRangeAddress(4 + count * rows, 4 + count * rows, 4, 5));

                count++;
            }



            const double width = 36.58;
            ws.SetColumnWidth(0, (int)(width * 180));
            ws.SetColumnWidth(1, (int)(width * 65));//73 * 256 / 7);
            ws.SetColumnWidth(2, (int)(width * 105));//111 * 256 / 7);
            ws.SetColumnWidth(3, (int)(width * 175));//189 * 256 / 7);
            ws.SetColumnWidth(4, (int)(width * 100));//107 * 256 / 7);
            ws.SetColumnWidth(5, (int)(width * 95));//101 * 256 / 7);
            /*
            ws.SetColumnWidth(0, (int)(20.63F * width));
            ws.SetColumnWidth(1, (int)(8.5F * width));//73 * 256 / 7);
            ws.SetColumnWidth(2, (int)(13.25F * width));//111 * 256 / 7);
            ws.SetColumnWidth(3, (int)(23.00F * width));//189 * 256 / 7);
            ws.SetColumnWidth(4, (int)(12.75F * width));//107 * 256 / 7);
            ws.SetColumnWidth(5, (int)(12.00F * width));//101 * 256 / 7);
            */

            FileStream file = new FileStream(@"out.xlsx", FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();
            Console.ReadKey();
        }
    }
}
