// See https://aka.ms/new-console-template for more information
Console.WriteLine("複数のワークシートを移動");

// ワークブックを開く
var workbook = new GrapeCity.Documents.Excel.Workbook();
workbook.Open("multiworksheet.xlsx");

// Sheet1とSheet2を末尾に移動
workbook.Worksheets[new string[] { "Sheet1", "Sheet2" }].Move();

//// Sheet1とSheet2をSheet3の後に移動
//workbook.Worksheets[new string[] { "Sheet1", "Sheet2" }].MoveAfter(workbook.Worksheets[2]);

//// Sheet3とSheet4をSheet1の前に移動
//workbook.Worksheets[new string[] { "Sheet3", "Sheet4" }].MoveBefore(workbook.Worksheets[0]);

// Excelファイルとして保存
workbook.Save("MoveMultipleWorksheets.xlsx");