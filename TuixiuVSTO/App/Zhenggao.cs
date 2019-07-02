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
using System.Globalization;

namespace TuixiuVSTO.App
{
    class Zhenggao : IDisposable
    {
        public static string workPath = @"D:\Library\Desktop\1\zhenggao\";
        public static string sumFileName = "正高女职工退休.xlsm";
        public static string templateFileName = "template3.docx";
        public static int KeyNum = 3;

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document doc;
        Excel.Range ranges;
        string PathHeader;

        public Zhenggao()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            excelApp.Quit();
            wordApp.Quit();
        }

        public void zhenggao2018()
        {

            if (!File.Exists(workPath + templateFileName))
            {
                MessageBox.Show("File cannot found");
                return;
            }

            try
            {
                thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            }
            catch
            {
                if (thisWorkBook != null)
                {
                    thisWorkBook.Close(false);
                }
                excelApp.Quit();
                MessageBox.Show("Excel Path not found");
                return;
            }

            try
            {
                thisWorkSheet = thisWorkBook.Worksheets["Sheet1"];
            }
            catch
            {
                thisWorkBook.Close(false);
                excelApp.Quit();
                MessageBox.Show("Sheet not found");
                return;
            }

            try
            {
                doc = wordApp.Documents.Open(workPath + templateFileName);
            }
            catch
            {
                if (doc != null)
                {
                    doc.Close();
                }
                MessageBox.Show("Word Path not found");
                return;
            }

            try
            {

                ranges = thisWorkSheet.UsedRange;

                /*
                string Name = "";
                string Year = "";
                string Month = "";
                */

                PathHeader = $@"{workPath}result\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }

                for (int i = 3; i <= ranges.Rows.Count; i++)
                {
                    foreach (Word.Variable var in doc.Variables)
                    {
                        //WriteLine(var.Name);
                        var.Value = "";
                    }


                    Dictionary<string, string> dict = new Dictionary<string, string>();

                    for (int j = 1; j <= KeyNum; j++)
                    {
                        string dictKey = ranges.Cells[2, j].Text;
                        string dictValue = ranges.Cells[i, j].Text;
                        WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                        dict.Add(dictKey, dictValue);
                    }

                    if (!(dict["Name"] == "" || dict["Name"] == null))
                    {
                        WriteLine("OK");
                        DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo
                        {
                            ShortDatePattern = "yyyy-MM-dd"
                        };
                        DateTime dt = Convert.ToDateTime(dict["Date1"], dtFormat);

                        doc.Variables.Add("Name", dict["Name"]);
                        doc.Variables.Add("Year", dt.Year);
                        doc.Variables.Add("Month", dt.Month);
                        //doc.Variables.Add("Date", dict["Date2"]);
                        doc.Variables.Add("Date", " ");


                        doc.Fields.Update();

                        doc.SaveAs2($@"{PathHeader}{dt.Year}-{dt.Month}-{dict["Name"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                    }

                }
            }
            finally
            {
                doc.Close();
                thisWorkBook.Close(false);
                excelApp.Quit();
            }

        }
    }
}
