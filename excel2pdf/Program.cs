using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace excel2pdf
{
    class Program
    {
        public static void setSheet2OnePage(Application excel, Workbook book)
        {
            excel.PrintCommunication = false;
            foreach (Worksheet sheet in book.Worksheets) {
                PageSetup setup = sheet.PageSetup;
                setup.Zoom = false;
                setup.FitToPagesWide = 1;
                setup.FitToPagesTall = false;
            }
            excel.PrintCommunication = true;
        }

        public static bool ExportWorkbookToPdf(string workbookPath, string outputPath)
        {
            // If either required string is null or empty, stop and bail out
            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(outputPath))
            {
                return false;
            }

            // Create COM Objects
            Application excelApplication;
            Workbook excelWorkbook;

            // Create new instance of Excel
            excelApplication = new Application();

            // Make the process invisible to the user
            excelApplication.ScreenUpdating = false;

            // Make the process silent
            excelApplication.DisplayAlerts = false;

            // Open the workbook that you wish to export to PDF
            excelWorkbook = excelApplication.Workbooks.Open(workbookPath);

            // If the workbook failed to open, stop, clean up, and bail out
            if (excelWorkbook == null)
            {
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;

                return false;
            }

            var exportSuccessful = true;
            try
            {
                setSheet2OnePage(excelApplication, excelWorkbook);
                // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (System.Exception ex)
            {
                // Mark the export as failed for the return value...
                exportSuccessful = false;

                // Do something with any exceptions here, if you wish...
                // MessageBox.Show...        
            }
            finally
            {
                // Close the workbook, quit the Excel, and clean up regardless of the results...
                excelWorkbook.Close();
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }

            // You can use the following method to automatically open the PDF after export if you wish
            // Make sure that the file actually exists first...
            // if (System.IO.File.Exists(outputPath))
            // {
            //     System.Diagnostics.Process.Start(outputPath);
            // }

            return exportSuccessful;
        }

        static int Main(string[] args)
        {
            var ap = new ArgumentParser();
            ap.Add('i', "input", OptionType.RequiredArgument, "input excel file path.");
            ap.Add('o', "output", OptionType.RequiredArgument, "output pdf file path.");
            ap.AddHelp();
            ap.Parse(args);
            String input = ap.Get("input");
            String output = ap.Get("output");
            bool ok = ExportWorkbookToPdf(input, output);
            return ok ? 0 : 1;
        }
    }
}
