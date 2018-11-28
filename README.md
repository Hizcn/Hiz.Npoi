# Hiz.Extended.Npoi
Npoi Extensions
[![NuGet](https://img.shields.io/nuget/v/Hiz.Npoi.svg)](https://www.nuget.org/packages/Hiz.Npoi/)


GetCellValue
```c#
cell.GetCellValue<DateTime>(DateTime.Now); // If CellType is Blank, return default value;
cell.GetCellValue<int?>(); // If CellType is Blank, return null;
cell.GetCellValue<double>(); // If CellType is string, return double.Parse(cell.StringCellValue);
```

Use Color
```c#
var style = workbook.CreateCellStyle();
style.SetFill(FillPattern.SolidForeground, "Red"); // ColorName
style.SetFill(FillPattern.SolidForeground, "#FF0000"); // HexString
style.SetFill(FillPattern.SolidForeground, 0xFFFF0000); // ArgbValue (has alpha)
style.SetFill(FillPattern.SolidForeground, IndexedColors.Red);
```

Simple apply border template, like office.
![Before](https://github.com/Hizcn/Hiz.Extended.Npoi/blob/master/images/01.png)
![After](https://github.com/Hizcn/Hiz.Extended.Npoi/blob/master/images/02.png)
```c#
var template = new BorderTemplate();
template.AddBorders(BorderTemplateEdges.AllAround, BorderStyle.Double, "Blue");
template.AddBorders(BorderTemplateEdges.AllInside, BorderStyle.Thin, 0xFFFF0000); // Red;
// x:columnIndex
// y:rowIndex
// width:columnCount
// height:rowCount
workbook.GetSheetAt(0).ApplyBorderTemplate(template, 1/*x*/, 2/*y*/, 6/*width*/, 9/*height*/);
```


Others
```c#
cell.GetCellTypeFinally(); // If CellType = Formula get CachedFormulaResultType;
cell.SetCellValue(object); // Support all scalar types, next version will support enum type.
cell.SetCellComment(string);

workbook.GetCellStyles();
workbook.GetFonts();
workbook.GetOrAddFont(...);
workbook.GetDefaultFont();
workbook.SetDefaultFont(...);
workbook.GetOrAddDataFormat();
workbook.GetSheets();
HSSFWorkbook.GetColorIndexed(NpoiColor); // If the palette is not standard, return the color index by ArgbValue.
XSSFWorkbook.GetColor(NpoiColor);

sheet.ApplyBorderTemplate(...);
sheet.AddRow();
sheet.AddRowRange(count);
sheet.InsertRow(index);
sheet.InsertRowRange(index, count);
sheet.RemoveRowRange(Func<IRow, bool> predicate, int start = 0, int count = -1, bool shift = true);
sheet.RemoveRowRangeWithEmpty();
sheet.GetColumnWidth(LengthUnit);
sheet.SetColumnWidth(...);
sheet.GetRows();
sheet.GetMaxRowCount();
sheet.GetFirstRowIndex(); // if sheet.PhysicalNumberOfRows=0 or 1, sheet.FirstRowNum both=0; and this method return -1 or 0;
sheet.GetLastRowIndex(); // if sheet.PhysicalNumberOfRows=0 or 1, sheet.LastRowNum both=0; and this method return -1 or 0;
sheet.GetOrAddRow();

row.GetCells();
row.GetOrAddCell();
row.IsEmpty();
row.RemoveAllCells();
row.RemoveAllCellsValues();
row.GetRowHeight(LengthUnit);
row.SetRowHeight(...);
row.GetCellValue<T>(default);

style.SetFill(...);
style.SetTextAlignment(...);
style.SetBorders(...);

```
