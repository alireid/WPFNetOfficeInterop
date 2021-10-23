using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using WPFNetOfficeInterop.Model;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace WPFNetOfficeInterop.Helpers
{
    public class ExcelHelper
    {
        public ExcelHelper()
        {
        }

        /// <summary>
        /// Output list of users as excel file using Interop
        /// </summary>
        /// <param name="users"></param>
        public static void Export(IList<User> users)
        {
            // Specify save path
            string excelFileName = Utilities.GetUniqueFilePath("C:\\temp\\excel.xls");

            // Setup new instance of excel app
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            int sheetId = 1;

            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Add
                (System.Reflection.Missing.Value,
                 xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count],
                 System.Reflection.Missing.Value,
                 System.Reflection.Missing.Value);

            // Remove first sheet
            if (sheetId == 1)
            {
                xlWorkBook.Worksheets[1].Delete();
            }

            // Recreate with sheet name
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetId);
            xlWorkSheet.Name = "Interop Excel Demo";

            // All white cells
            var xlRange = xlWorkSheet.get_Range("A1", "Z40");
            xlRange.Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;
            xlRange.Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlNone;
            xlRange.Interior.TintAndShade = 0;
            xlRange.Interior.PatternTintAndShade = 0;

            // Header cells and text
            xlWorkSheet.Cells[1, 1] = "User List";
            xlWorkSheet.Cells[2, 1] = "Demonstration of office interop using .NET WPF C# - Alasdair Reid 2021";
            xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[2, 1].EntireRow.Font.Bold = true;
            xlWorkSheet.get_Range("A1", "A3").Cells.Font.Size = 14;
            xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 10]].Merge();
            xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 10]].Merge();

            // Table data headings
            xlWorkSheet.Cells[4, 1] = "First name";
            xlWorkSheet.Cells[4, 2] = "Last name";
            xlWorkSheet.Cells[4, 3] = "City";
            xlWorkSheet.Cells[4, 4] = "Post code";

            // Loop over data for user data
            int row = 5;
            foreach (User u in users)
            {
                for (int i = 1; i <= 4; i++)
                {
                    string fieldText = string.Empty;
                    switch (i)
                    {
                        case 1:
                            fieldText = u.FirstName;
                            break;
                        case 2:
                            fieldText = u.LastName;
                            break;
                        case 3:
                            fieldText = u.City;
                            break;
                        case 4:
                            fieldText = u.Postcode;
                            break;
                    }
                    xlWorkSheet.Cells[row, i] = fieldText;
                }
                row++;
            }

            // Autofit
            for (int i = 1; i <= users.Count+5; i++)
            {
                xlWorkSheet.Columns[i].EntireColumn.AutoFit();
            }

            // Save file
            xlWorkBook.SaveAs(excelFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
        }
    }
}
