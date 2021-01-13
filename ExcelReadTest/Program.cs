
using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;


 
namespace ExcelReadTest
{
    class Program
    {
        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook = ExcelFile.Load(@"Test.xlsx");

            var worksheet = workbook.Worksheets[0];

            List<Model> testModel = new List<Model>();

            //create Datatable 
            var dataTable = worksheet.CreateDataTable(new CreateDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0,
                NumberOfColumns = 2,
                NumberOfRows = worksheet.Rows.Count,
                Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
            });

            // write datatable content
            var sb = new StringBuilder();
            sb.AppendLine("Content: ");
            foreach (DataRow row in dataTable.Rows)
            {
                //sb.AppendFormat($"{row[0]}\t{row[1]}");
                //sb.AppendLine();

                testModel.Add(new Model { FirstCol = row[0].ToString(), SecondCol = row[1].ToString() });
            }

            //Console.WriteLine(sb.ToString());

            foreach (var test in testModel)
            {
                Console.WriteLine($"First = {test.FirstCol}\t Second = {test.SecondCol}");
            }
            Console.ReadKey();
        }

        class Model
        {
            public string FirstCol { get; set; }
            public string SecondCol { get; set; }
        }
    }
}
