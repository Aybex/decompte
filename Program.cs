// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;

Console.WriteLine("Hello, World!");

using var workbook = new XLWorkbook("Decompte.xlsx");

var worksheet = workbook.Worksheets.Worksheet("Decompte");

var marche = worksheet.Cell("C6").Value;

//worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
Console.WriteLine(marche);

//workbook.SaveAs("HelloWorld.xlsx");
