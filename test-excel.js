// excelオブジェクトの作成と表示する。
var excel = WScript.CreateObject("Excel.Application");
excel.Visible = true; 
// 新規ブックを作成する。
var book = excel.Workbooks.Add();
var sheet = book.ActiveSheet;
// セルを指定して値を入れる。
sheet.Cells(1,1).Value = "hoge";
