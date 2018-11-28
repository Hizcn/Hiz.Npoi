using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Hiz.Npoi;
using NPOI.SS.UserModel;
using System.IO;

namespace Sample.ConsoleApp
{
    class Sample
    {
        public static void TestBorderTemplate()
        {
            var template = new BorderTemplate();
            template.AddBorders(BorderTemplateEdges.AllAround, BorderStyle.Double, "Blue");
            template.AddBorders(BorderTemplateEdges.AllInside, BorderStyle.Thin, 0xFFFF0000); // Red;

            IWorkbook workbook;
            using (var file = File.OpenRead("test.xlsx"))
            {
                workbook = WorkbookFactory.Create(file);
            }
            workbook.GetSheetAt(0).ApplyBorderTemplate(template, 1, 2, 6, 9);
            workbook.Write("test.xlsx");
        }

        public static void TestColor()
        {
            IWorkbook workbook = new XSSFWorkbook();
            var style = workbook.CreateCellStyle();
            style.SetFill(FillPattern.SolidForeground, "Red"); // ColorName
            style.SetFill(FillPattern.SolidForeground, "#FF0000"); // HexString
            style.SetFill(FillPattern.SolidForeground, 0xFFFF0000); // ArgbValue (has alpha)
            style.SetFill(FillPattern.SolidForeground, IndexedColors.Red); // IndexedColors
        }
    }
}
