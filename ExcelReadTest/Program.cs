
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelReadTest
{
    class Program
    {
        static List<Model> testModel = new List<Model>();

        static void Main(string[] args)
        {

            FileInfo file = new FileInfo("test.xlsx");

            ExcelPackage.LicenseContext = LicenseContext.Commercial; // set the lineces

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int colCount = worksheet.Dimension.End.Column; // get column count
                int rowCount = worksheet.Dimension.End.Row; // get row count

                for (int r = 2; r <= rowCount; r++) // start at 2 to skip the header.
                {
                    List<string> col = new List<string>();
                    for (int c = 1; c <= colCount; c++)
                    {
                        //Console.WriteLine($"Row:{r} Col:{c} Value: {worksheet.Cells[r, c].Value?.ToString()}");

                        col.Add(worksheet.Cells[r, c].Value?.ToString());
                    }

                    testModel.Add(new Model
                    {
                        ColA = col[0],
                        ColB = col[1]
                    });
                }
            }

            foreach (var t in testModel)
            {
                Console.WriteLine($"{t.ColA}\t{t.ColB}");
            }

            Console.ReadKey();
        }

    }

    class Model
    {
        public string ColA { get; set; }
        public string ColB { get; set; }
    }
}
