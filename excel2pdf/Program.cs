using Microsoft.Office.Interop.Excel;
using System;
using System.Security;
using Application = Microsoft.Office.Interop.Excel.Application;


namespace excel2pdf
{
    class Program
    {
        [SecurityCritical]
        public static string ExportWorkbookToPdf(string workbookPath, string outputPath)
        {
            // 如果所需字符串为 null 或为空，则停止并退出
            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(outputPath))
            {
                return "字符串为空";
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

            // 打开您要导出为 PDF 的工作簿
            excelWorkbook = excelApplication.Workbooks.Open(workbookPath);

            // 如果工作簿无法打开、停止、清理和退出
            if (excelWorkbook == null)
            {
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;

                return "工作簿无法打开、停止、清理和退出";
            }

            var exportSuccessful = "OK";
            try
            {
                //excelApplication.PrintCommunication = false;
                foreach (Worksheet sheet in excelWorkbook.Worksheets)
                {
                    PageSetup setup = sheet.PageSetup;
                    setup.Zoom = false;
                    setup.FitToPagesWide = 1;
                    setup.FitToPagesTall = false;
                }
                //excel Application.Print 通信 = true;
                // 调用Excel原生导出函数（在Office 2007和Office 2010中有效，AFAIK）
                excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (System.Exception ex)
            {
                // Mark the export as failed for the return value...
                exportSuccessful = ex.Message;

                // Do something with any exceptions here, if you wish...
                // MessageBox.Show(ex.Message);
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
            Environment.SetEnvironmentVariable("COMPlus_legacyCorruptedStateExceptionsPolicy", "1");
            var ap = new ArgumentParser();
            ap.Add('i', "input", OptionType.RequiredArgument, "input excel file path.");
            ap.Add('o', "output", OptionType.RequiredArgument, "output pdf file path.");
            ap.AddHelp();
            ap.Parse(args);
            var input = ap.Get("input") ?? $@"{AppDomain.CurrentDomain.BaseDirectory}Cliente_11_9页长单据.xlsx";
            var output = ap.Get("output") ?? $@"{AppDomain.CurrentDomain.BaseDirectory}test.pdf";
            if (!System.IO.File.Exists(input))
            {
                Console.WriteLine("xlsx 文件不存在或无法访问");
                return 0;
            }
            if (System.IO.File.Exists(output))
            {
                System.IO.File.Delete(output);
            }
            var res=ExportWorkbookToPdf(input, output);
            Console.WriteLine(res);
            return res == "OK"?1:0;
        }
    }
}
