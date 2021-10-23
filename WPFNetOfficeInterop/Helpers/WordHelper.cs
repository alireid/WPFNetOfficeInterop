using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using WPFNetOfficeInterop.Model;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WPFNetOfficeInterop.Helpers
{
    public class WordHelper
    {

        public WordHelper()
        {
        }

        /// <summary>
        /// Output list of users as word file using Interop
        /// </summary>
        /// <param name="users"></param>
        public static void Export(IList<User> users)
        {
            // Specify save path
            string wordFileName = Utilities.GetUniqueFilePath("C:\\temp\\word.docx");

            //Create an instance for word
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            //Set animation status for word application
            wordApp.ShowAnimation = false;

            //Set to open word for the user
            wordApp.Visible = true;

            //Create a missing variable for missing value
            object missing = System.Reflection.Missing.Value;

            //Create a new document
            Microsoft.Office.Interop.Word.Document document = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            //adding text to document
            document.Content.SetRange(0, 0);

            //Add paragraph with Heading 1 style
            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading1 = "Heading 1";
            para1.Range.set_Style(ref styleHeading1);
            para1.Range.Text = "User List";
            para1.Range.InsertParagraphAfter();

            //Add paragraph with Heading 2 style
            Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading2 = "Heading 2";
            para1.Range.set_Style(ref styleHeading2);
            para1.Range.Text = "Demonstration of office interop using .NET WPF C# - Alasdair Reid 2021";
            para1.Range.InsertParagraphAfter();

            //Add paragraph with for the table data
            Microsoft.Office.Interop.Word.Paragraph para3 = document.Content.Paragraphs.Add(ref missing);
            Microsoft.Office.Interop.Word.Table table = document.Tables.Add(para3.Range, users.Count+1, 4, ref missing, ref missing);
            table.Borders.Enable = 1;

            // Header text
            table.Cell(1, 1).Range.Text = "First name";
            table.Cell(1, 2).Range.Text = "Last name";
            table.Cell(1, 3).Range.Text = "City";
            table.Cell(1, 4).Range.Text = "Post code";

            // Loop data
            int row = 2;
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
                    table.Cell(row, i).Range.Text = fieldText;
                }
                row++;
            }
            document.SaveAs2(wordFileName);
        }

    }
}
