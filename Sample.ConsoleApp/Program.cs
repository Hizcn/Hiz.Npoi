using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Hiz.Npoi;
using NPOI.SS.UserModel;
using System.Diagnostics;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;
using System.IO;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;

namespace Sample.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Sample.TestBorderTemplate();

            TestColor.TestColorOperators();
            //Test7();


            Console.WriteLine("OK");
            Console.ReadLine();
        }
        static void Test7()
        {
            var workbook1x = CreateWorkbook(true);
            var workbook1s = CreateWorkbook(false);

            var bt = new BorderTemplate();
            bt.AddBorders(BorderTemplateEdges.AllInside, BorderStyle.Thin, 0xFFFF0000);

            workbook1x.GetSheetAt(0).ApplyBorderTemplate(bt, 2, 2, 10, 10);
            workbook1s.GetSheetAt(0).ApplyBorderTemplate(bt, 2, 2, 10, 10);

            workbook1s.Write("test7.xls");
            workbook1x.Write("test7.xlsx");
        }
        static void Test6()
        {

            var workbook1x = CreateWorkbook(true);
            var workbook1s = CreateWorkbook(false);

            //var fonts1x = workbook1x.GetFonts().ToArray();
            //var font1x = workbook1x.CreateFont();
            //var font1s = workbook1s.CreateFont();


            var styles1x = workbook1x.GetCellStyles().ToArray();
            var styles1s = workbook1s.GetCellStyles().ToArray();

            var style1x = (XSSFCellStyle)workbook1x.CreateCellStyle();
            var style1s = (HSSFCellStyle)workbook1s.CreateCellStyle();
            //style1s.BorderDiagonalColor = IndexedColors.Automatic.Index;
            //style1s.BorderDiagonalColor = 0;
            //var font1x= workbook1x.GetDefaultFont();
            //var font1s = workbook1s.GetDefaultFont();

            style1x.SetFill(FillPattern.SolidForeground, "Red", IndexedColors.Blue);
            style1s.SetFill(FillPattern.SolidForeground, "Red", IndexedColors.Blue);

            var cell0x = workbook1x.GetSheetAt(0).GetOrAddRow(0).GetOrAddCell(0);
            //cell0x.CellStyle = style1x;
            var cell0s = workbook1s.GetSheetAt(0).GetOrAddRow(0).GetOrAddCell(0);
            //cell0s.CellStyle = style1s;


            workbook1x.Write("test6.xlsx");
            workbook1s.Write("test6.xls");
        }



        static void Test5()
        {
            //NpoiColor.asdasd("");

            //var v6 = IndexedColors.ValueOf(27);
            //NpoiColor.SetupMappings();

            var v21 = new byte[] { 0x20, 0x30, 0x40 };
            var v22 = new byte[] { 0x10, 0x20, 0x30, 0x40 };

            var v23 = BitConverter.ToInt32(v22, 0).ToString("X8");
            var v24 = BitConverter.ToInt32(v22.Reverse().ToArray(), 0).ToString("X8");
            var v25 = BitConverter.ToInt32(v21.Reverse().ToArray(), 0).ToString("X8");

            var workbook1x = (XSSFWorkbook)CreateWorkbook(true);

            var sheet0 = workbook1x.CreateSheet();
            var row0 = sheet0.GetOrAddRow(0);
            var r0cell0 = row0.GetOrAddCell(0);

            var c1 = new XSSFColor(IndexedColors.Blue.RGB);
            var font = (XSSFFont)workbook1x.GetDefaultFont();
            font.SetColor(c1);

            var style = workbook1x.CreateCellStyle();
            style.SetFont(font);

            r0cell0.CellStyle = style;
            r0cell0.SetCellValue("Test");

            var v1 = workbook1x.GetTheme();
            var v2 = workbook1x.GetStylesSource();
            var v3 = v2.GetTheme();

            //var theme = new NPOI.XSSF.Model.ThemesTable();
            //v2.SetTheme(theme);

            //var v21 = workbook1x.GetTheme();
            //var v22 = workbook1x.GetStylesSource();
            //var v23 = v2.GetTheme();

            //workbook1x.Write("test2.xlsx");
        }

        static void Test1()
        {
            // 全部移除
            var workbook1x = CreateWorkbook(true);
            var workbook1s = CreateWorkbook(false);
            var count1x = workbook1x.GetSheetAt(0).RemoveRowRangeWithEmpty();
            var count1s = workbook1s.GetSheetAt(0).RemoveRowRangeWithEmpty();
            var result1 = new HashSet<int>(new[] { 0, 1, 2, 3, 4, 5 });
            var rows1x = workbook1x.GetSheetAt(0).GetRows().Select(i => i.RowNum).ToArray();
            var rows1s = workbook1s.GetSheetAt(0).GetRows().Select(i => i.RowNum).ToArray();
            //Debug.Assert(count1x == 20 - 6);
            //Debug.Assert(count1s == 20 - 6);
            Debug.Assert(result1.SetEquals(rows1x));
            Debug.Assert(result1.SetEquals(rows1s));


            //var workbook222x = CreateWorkbook(false);

            // 移除上部
            var workbook2x = CreateWorkbook(true);
            var workbook2s = CreateWorkbook(false);
            var count2x = workbook2x.GetSheetAt(0).RemoveRowRangeWithEmpty(2, 7); // 2/3/4/5/6/7/8 (其中4/7有值 其它 5 行空白)
            var count2s = workbook2s.GetSheetAt(0).RemoveRowRangeWithEmpty(2, 7);
            var result2 = new HashSet<int>(new[] { 2, 3, 5, 8, 11, 14 });
            var rows2x = workbook2x.GetSheetAt(0).GetRows().Select(i => i.RowNum).ToArray();
            var rows2s = workbook2s.GetSheetAt(0).GetRows().Select(i => i.RowNum).OrderBy(i => i).ToArray();
            //Debug.Assert(count2x == 5);
            //Debug.Assert(count2s == 5);
            Debug.Assert(result2.SetEquals(rows2x));
            Debug.Assert(result2.SetEquals(rows2s));

            // 移除下部
            var workbook3x = CreateWorkbook(true);
            var workbook3s = CreateWorkbook(false);
            var count3x = workbook3x.GetSheetAt(0).RemoveRowRangeWithEmpty(12, 10); // 12/13/14/15/16/17/18/19/20/21 (其中13/16/19有值 其它 5 行空白, 20/21 行超出 LastRowIndex)
            var count3s = workbook3s.GetSheetAt(0).RemoveRowRangeWithEmpty(12, 10);
            var result3 = new HashSet<int>(new[] { 4, 7, 10, 12, 13, 14 });
            var rows3x = workbook3x.GetSheetAt(0).GetRows().Select(i => i.RowNum).ToArray();
            var rows3s = workbook3s.GetSheetAt(0).GetRows().Select(i => i.RowNum).OrderBy(i => i).ToArray();
            //Debug.Assert(count3x == 5);
            //Debug.Assert(count3s == 5);
            Debug.Assert(result3.SetEquals(rows3x));
            Debug.Assert(result3.SetEquals(rows3s));

            // 移除中部
            var workbook4x = CreateWorkbook(true);
            var workbook4s = CreateWorkbook(false);
            var count4x = workbook4x.GetSheetAt(0).RemoveRowRangeWithEmpty(6, 9); // 6/7/8/9/10/11/12/13/14 (其中7/10/13有值 其它 6 行空白)
            var count4s = workbook4s.GetSheetAt(0).RemoveRowRangeWithEmpty(6, 9);
            var result4 = new HashSet<int>(new[] { 4, 6, 7, 8, 10, 13 });
            var rows4x = workbook4x.GetSheetAt(0).GetRows().Select(i => i.RowNum).ToArray();
            var rows4s = workbook4s.GetSheetAt(0).GetRows().Select(i => i.RowNum).OrderBy(i => i).ToArray();
            //Debug.Assert(count4x == 6);
            //Debug.Assert(count4s == 6);
            Debug.Assert(result4.SetEquals(rows4x));
            Debug.Assert(result4.SetEquals(rows4s));
        }

        static void Test2()
        {
            var workbook1x = CreateWorkbook(true);
            var workbook1s = CreateWorkbook(false);

            //var fs1x = workbook1x.GetFormats(true).ToArray();
            //var fs1s = workbook1s.GetFormats(true).OrderBy(i => i.Index).ToArray();
            //Test3("x", fs1x);
            //Test3("s", fs1s);
        }
        static void Test4()
        {
            //var path1 = @"wps1.xlsx";
            //var workbook1x = WorkbookFactory.Create(path1);
            //var fs1x = workbook1x.GetFormats(true).ToArray();
            //Test3("x", fs1x);
            var path2 = @"wps2.xlsx";

            using (var file = File.OpenRead(path2))
            {
                var w = WorkbookFactory.Create(file);
                //var sss= w.GetFormats().ToArray();
            }

            //    var workbook2x = WorkbookFactory.Create(path2);
            //var fs2x = workbook2x.GetFormats(true).ToArray();
            //Test3("x", fs2x);

            //var path3 = @"wps3.xls";
            //var workbook3x = WorkbookFactory.Create(path3);
            //var fs3x = workbook3x.GetFormats(true).ToArray();
            //Test3("x", fs3x);

            //var workboos4x = CreateWorkbook(true);
            //var fs4x = workboos4x.GetFormats(true).ToArray();
            //Test3("x", fs4x);

            //var workboos5x = CreateWorkbook(false);
            //var fs5x = workboos5x.GetFormats(true).ToArray();
            //Test3("x", fs5x);
        }

        //static void Test3(string name,IEnumerable<ExcelDataFormat> formats)
        //{
        //    Console.WriteLine(name);
        //    foreach (var f in formats)
        //    {
        //        Console.WriteLine("0x{0:X2}: {1}/{2}", new object[] { f.Index, f.Source, f.Format });
        //    }
        //}

        static IWorkbook CreateWorkbook(bool xlsx, string path = null)
        {
            var workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet0 = workbook.CreateSheet();

            //var row1 = sheet0.CreateRow(1);
            //var row3 = sheet0.CreateRow(3);
            //var row5 = sheet0.CreateRow(5);
            //var rows = new[] { row1, row3, row5 };
            //foreach(var r in rows)
            //{
            //}

            var indexes = new int[] { 4, 7, 10, 13, 16, 19 };
            foreach (var i in indexes)
            {
                var row = sheet0.CreateRow(i);
                row.CreateCell(0).SetCellValue("Test" + i);
            }

            //workbook.GetOrAddFormat("######");

            //sheet0.ShiftRows(7, 7, -3);

            //var sss = sheet0.GetRows().Select(i => i.RowNum).OrderBy(i => i).ToArray();
            //var v1 = sheet0.GetRow(4).GetCellValue(0, string.Empty);
            //var ss2s = sheet0.GetRows().Where(r => r.PhysicalNumberOfCells == 0).Select(i => i.RowNum).OrderBy(i => i).ToArray();

            if (!string.IsNullOrEmpty(path))
                workbook.Write(path);
            return workbook;
        }
    }
}
