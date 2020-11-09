using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using static System.Diagnostics.Debug;

namespace TuixiuVSTO.App
{
    class _MergeExcelSheet
    {

        public void mergeMultiSheetToOneWorksheet()
        {
            string workPath = @"D:\1\VSTO_MergeExcelSheet\";
            string mergeFileName = "combine.xlsx";

            string PathHeader = $@"{workPath}result\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            if (!Directory.Exists(PathHeader))
            {
                Directory.CreateDirectory(PathHeader);
            }


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook combineWorkBook = excelApp.Workbooks.Add();
            Excel.Worksheet combineWorkSheet = combineWorkBook.Worksheets.Add();

            WriteLine(workPath);

            DirectoryInfo workDir = new DirectoryInfo(workPath);
            FileInfo[] files = workDir.GetFiles("*.xls*");

            foreach (FileInfo file in files)
            {
                if (FileAttributes.Hidden != (file.Attributes & FileAttributes.Hidden) && file.Name != mergeFileName)
                {
                    WriteLine(file.FullName);

                    Excel.Workbook workbook = excelApp.Workbooks.Open(file.FullName);

                    Excel.Worksheet copyWorkSheet = workbook.Worksheets[1];

                    if (copyWorkSheet != null)
                    {
                        copyWorkSheet.UsedRange.Copy();
                        combineWorkSheet.Paste();

                    }
                    else
                    {
                        WriteLine($"Sheet1 Not Found: ${file.Name}");
                    }

                    workbook.Close(false);

                }
            }

            combineWorkBook.SaveAs(Filename: $@"{workPath}{mergeFileName}");
            combineWorkBook.Close();
        }
    }
}
