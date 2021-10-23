using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using WPFNetOfficeInterop.Model;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace WPFNetOfficeInterop.Helpers
{
    public class PowerPointHelper
    {
        public PowerPointHelper()
        {
        }

        /// <summary>
        /// Output list of users as powerpoint file using Interop
        /// </summary>
        /// <param name="users"></param>
        public static void Export(IList<User> users)
        {
            // Specify save path
            string powerpointFileName = Utilities.GetUniqueFilePath("C:\\temp\\powerpoint.pptx");

            //Create an instance for powerpoint
            Application pptApplication = new Application();
            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            Microsoft.Office.Interop.PowerPoint.TextRange objText;
            Microsoft.Office.Interop.PowerPoint.Shape objShape;

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Create new Slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            // Add title / heading details
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = "User List";
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = "List of users created using .NET WPF - Alasdair Reid 2021";

            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Demonstration of office interop using .NET WPF C# - Alasdair Reid 2021";

            // Add table
            int iRow;
            int iColumn;

            int iRows = users.Count;
            int iColumns = 3;

            objShape = slide.Shapes.AddTable(iRows, iColumns, shape.Left, 110, shape.Width, 120);

            for (iRow = 1; iRow <= objShape.Table.Rows.Count; iRow++)
            {
                for (iColumn = 1; iColumn <= objShape.Table.Columns.Count; iColumn++)
                {
                    string fieldText = string.Empty;

                    switch (iColumn)
                    {
                        case 1:
                            fieldText = users[iRow-1].FirstName;
                            break;
                        case 2:
                            fieldText = users[iRow-1].LastName;
                            break;
                        case 3:
                            fieldText = users[iRow-1].City;
                            break;
                        case 4:
                            fieldText = users[iRow-1].Postcode;
                            break;

                    }

                    objShape.Table.Cell(iRow, iColumn).Shape.TextFrame.TextRange.Text = fieldText;
                    objShape.Table.Cell(iRow, iColumn).Shape.TextFrame.TextRange.Font.Name = "Verdana";
                    objShape.Table.Cell(iRow, iColumn).Shape.TextFrame.TextRange.Font.Size = 8;
                }
            }

            // Save file and open powerpoint
            pptPresentation.SaveAs(powerpointFileName, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptApplication.Presentations.Application.Activate();
            pptApplication.Activate();
        }

    }
}
