// ReadExcelSample.csx — đọc file Excel mẫu in ra console
using System;
using System.IO;
using ClosedXML.Excel;

var path = @"D:\Work\Freelance\TextInputter\data\27-02-2026\excel\Hàng lấy\NGAY 27-2.xlsx";
using var wb = new XLWorkbook(path);

foreach (var ws in wb.Worksheets)
{
    Console.WriteLine($"\n{'='*60}");
    Console.WriteLine($"SHEET: {ws.Name}");
    Console.WriteLine($"{'='*60}");
    var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
    var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
    Console.WriteLine($"Size: {lastRow} rows x {lastCol} cols");
    Console.WriteLine();

    for (int r = 1; r <= lastRow; r++)
    {
        var parts = new System.Collections.Generic.List<string>();
        for (int c = 1; c <= lastCol; c++)
        {
            var cell = ws.Cell(r, c);
            var val = cell.Value.ToString();
            if (!string.IsNullOrEmpty(cell.FormulaA1))
                val = $"[={cell.FormulaA1}]={val}";
            parts.Add($"C{c}:{val}");
        }
        Console.WriteLine($"R{r}: {string.Join(" | ", parts)}");
    }
}
