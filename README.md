# ExcelReadTest_Console

1. First is to install the latest version of *EPPlus* in Nugetpackages

1.  Initilize where the Excel file is located at #####(My test spreadsheet is located in the project itself)

```
FileInfo file = new FileInfo("test.xlsx");
```

1. Set the licensce to be used by the program 
   - Set a Commercial license to use an existing one ``` ExcelPackage.LicenseContext = LicenseContext.Commercial; ```
 
> Set a Non-Commercial license to use an existing one ``` ExcelPackage.LicenseContext = LicenseContext.NonCommercial; ```
 
 
1. Set up a Model based off of your spread sheet
```
 class Model
{
    public string ColA { get; set; }
    public string ColB { get; set; }
}
```

1. Setup the *Using* feild to connect to the spreedsheet

```
using (ExcelPackage package = new ExcelPackage(file){}
```

1. Setup which page you want to copy from
```
ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
```

1. Get the count of how the Rows and Columns are in the spreadsheet

```
int colCount = worksheet.Dimension.End.Column; // get column count
int rowCount = worksheet.Dimension.End.Row; // get row count
```

1. Add the contents of the spreedsheet to a list of the model, 
```
List<Model> testModel = new List<Model>();

for (int r = 2; r <= rowCount; r++) // start at 2 to skip the header.
{
  List<string> col = new List<string>();
  for (int c = 1; c <= colCount; c++)
  {
      col.Add(worksheet.Cells[r, c].Value?.ToString());
  }

  testModel.Add(new Model
  {
      ColA = col[0],
      ColB = col[1]
  });
}
```
