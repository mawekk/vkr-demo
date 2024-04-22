using static DocUtils.Xlsx;
using Microsoft.FSharp.Core;

Spreadsheet spreadsheet = Spreadsheet.New(FSharpOption<string>.None);

var sheet = spreadsheet.Sheets().First();
var data = new string[] { "hello", "world", "!" };

for (int i = 0; i < 5; i++)
{
    sheet.WriteRow(data);
}

await spreadsheet.SaveTo("test.xlsx");
