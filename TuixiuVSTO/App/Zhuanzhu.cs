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

namespace TuixiuVSTO.App
{
    class Zhuanzhu
    {

        public void zhuanzhu()
        {
            int NArow = 0;
            string workPath = @"D:\VS2017\Projects\ExcelWorkbook1\ExcelWorkbook1\bin\Debug\test\";
            string sumFileName = "test.xlsx";

            Excel.Application excelApp = new Excel.Application();
            //Excel.Workbook thisWorkBook = excelApp.ActiveWorkbook;
            //Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(excelApp.ActiveWorkbook.Path + "\\a.xlsx");
            Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            Excel.Worksheet thisWorkSheet = thisWorkBook.Worksheets["Sheet0"];

            //workPath = thisWorkBook.Path;
            WriteLine(workPath);

            DirectoryInfo workDir = new DirectoryInfo(workPath);
            FileInfo[] files = workDir.GetFiles("*.xls*");

            foreach (FileInfo file in files)
            {
                if (FileAttributes.Hidden != (file.Attributes & FileAttributes.Hidden) && file.Name != sumFileName)
                {
                    WriteLine(file.FullName);

                    Excel.Workbook workbook = excelApp.Workbooks.Open(file.FullName);

                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        WriteLine(sheet.Name);
                        if (sheet.Name != "Sheet0")
                        {
                            int i = 0;
                            NArow = thisWorkSheet.Range["A1"].CurrentRegion.Rows.Count + 1;

                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B2"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["D2"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B3"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B4"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B5"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B6"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["D6"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B7"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["D7"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B8"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["D8"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B9"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["D9"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B10"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B11"].Value;
                            thisWorkSheet.Cells[NArow, ++i] = sheet.Range["B12"].Value;

                        }
                    }

                    workbook.Close(false);
                }
            }

            //thisWorkBook.Save();

        }

        public void zhuanzhu2()
        {
            int NArow = 0;
            string workPath = @"D:\VS2017\Projects\ExcelWorkbook1\ExcelWorkbook1\bin\Debug\test\";
            string sumFileName = "test.xlsx";

            Excel.Application excelApp = new Excel.Application();
            //Excel.Workbook thisWorkBook = excelApp.ActiveWorkbook;
            //Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(excelApp.ActiveWorkbook.Path + "\\a.xlsx");
            Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            Excel.Worksheet thisWorkSheet = thisWorkBook.Worksheets["Sheet0"];

            //workPath = thisWorkBook.Path;
            WriteLine(workPath);

            DirectoryInfo workDir = new DirectoryInfo(workPath);
            FileInfo[] files = workDir.GetFiles("*.xls*");

            int NAcol = thisWorkSheet.Range["A1"].CurrentRegion.Columns.Count;

            foreach (FileInfo file in files)
            {
                if (FileAttributes.Hidden != (file.Attributes & FileAttributes.Hidden) && file.Name != sumFileName)
                {
                    WriteLine(file.FullName);

                    Excel.Workbook workbook = excelApp.Workbooks.Open(file.FullName);

                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        WriteLine(sheet.Name);
                        if (sheet.Name != "Sheet0")
                        {
                            //int i = 0;
                            NArow = thisWorkSheet.Range["A1"].CurrentRegion.Rows.Count + 1;

                            Excel.Range ranges = sheet.Range["A1"].CurrentRegion;

                            for (int i = 1; i <= NAcol; i++)
                            {
                                string key = thisWorkSheet.Cells[1, i].Value;
                                WriteLine($"key: {key}");

                                foreach (Excel.Range range in ranges)
                                {
                                    string key2 = range.Value as string;
                                    if (key2 == key)
                                    {
                                        WriteLine($"key2: {key2}");
                                        var perVar = sheet.Cells[range.Row, range.Column + 1].Value;

                                        if (perVar != null)
                                        {
                                            /* You should use {as string} only if you are sure of the type of Value.
                                             * To get a string from any kind of value in the cell, you can use the ToString() method.*/
                                            string value = perVar.ToString();

                                            WriteLine($"value: {value}");
                                            thisWorkSheet.Cells[NArow, i] = value;
                                        }

                                    }
                                }
                            }
                        }
                    }

                    workbook.Close(false);
                }
            }

            //thisWorkBook.Save();

        }

    }
}
