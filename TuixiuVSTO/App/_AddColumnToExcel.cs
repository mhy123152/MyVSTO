using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using VB = Microsoft.VisualBasic;
using System.Diagnostics;
using System.IO;
using static System.Diagnostics.Debug;
using System.Collections.Generic;
using TuixiuVSTO.Modules;
using System.Linq;
using Microsoft.Office.Interop.Word;
using TuixiuVSTO.App;
using Application = System.Windows.Forms.Application;
using Microsoft.Office.Core;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace TuixiuVSTO.App
{
    class _AddColumnToExcel : IDisposable
    {
        #region 需要修改的变量
        public static string columnName = "Date";
        #endregion

        #region 固定变量
        public static string workPath = @"D:\1\VSTO_AddColumnToExcel\";

        #endregion

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Range ranges;

        string PathHeader;

        public _AddColumnToExcel()
        {
            excelApp = new Excel.Application();
        }

        public void Dispose()
        {
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            excelApp = null;
        }


        public void run()
        {
            PathHeader = $@"{workPath}result\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            if (!Directory.Exists(PathHeader))
            {
                Directory.CreateDirectory(PathHeader);

            }

            DirectoryInfo workPathDirInfo = new DirectoryInfo(workPath);

            try
            {
                foreach (FileInfo fileInfo in workPathDirInfo.GetFiles())//遍历文件夹下的每个文件
                {
                    WriteLine(fileInfo.Name);
                    doBatch(fileInfo);
                }

            }
            finally
            {
                Dispose();
            }



        }

        private void doBatch(FileInfo fileInfo)
        {
            string filePath = $@"{workPath}{fileInfo.Name}";

            if (!File.Exists(filePath))
            {
                MessageBox.Show("File cannot found");
                Application.Exit();
            }

            thisWorkBook = excelApp.Workbooks.Open(filePath);

            thisWorkSheet = thisWorkBook.Worksheets[1];

            //WriteLine(thisWorkSheet.Name);

            string filename2 = System.IO.Path.GetFileNameWithoutExtension(fileInfo.Name);

            string addColString = "";
            if (filename2.Contains("-13"))
            {
                addColString = filename2.Replace("HT", "");
            }
            else
            {
                addColString = $@"{filename2.Replace("HT", "")}-01";
            }



            Excel.Range oRng = thisWorkSheet.Range["A1"];
            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

            oRng = thisWorkSheet.UsedRange.Columns["A:A", Type.Missing];
            oRng.Value2 = addColString;

            thisWorkSheet.Cells[1, 1] = columnName;

            thisWorkBook.SaveAs(Filename: $@"{PathHeader}{filename2}.xlsx", FileFormat: 51);

            thisWorkBook.Close(false);
        }
    }
}
