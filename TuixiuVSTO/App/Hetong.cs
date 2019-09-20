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
    class Hetong : IDisposable
    {
        public static string workPath = @"D:\Library\Desktop\1\hetong\";
        public static string sumFileName = "Data.xlsx";
        public static string templateFileName = "事业单位聘用合同（人事代理续签模板）.docx";
        public static string templateFileName2 = "事业单位聘用合同（合同制续签模板）.docx";
        public static int KeyNum = 14;

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Range ranges;
        string PathHeader;

        public Hetong()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            excelApp.Quit();
            wordApp.Quit();
        }

        public void hetong()
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


            Word.Document doc1 = null;
            Word.Document doc2 = null;
            try
            {
                doc1 = wordApp.Documents.Open(workPath + templateFileName);
                doc2 = wordApp.Documents.Open(workPath + templateFileName2);
            }
            catch
            {
                if (doc1 != null)
                {
                    doc1.Close();
                }
                if (doc2 != null)
                {
                    doc2.Close();
                }
                MessageBox.Show("Word Path not found");
                return;
            }

            try
            {

                ranges = thisWorkSheet.UsedRange;

                PathHeader = $@"{workPath}result\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }

                for (int i = 3; i <= ranges.Rows.Count; i++)
                {

                    Dictionary<string, string> dict = new Dictionary<string, string>();

                    for (int j = 1; j <= KeyNum; j++)
                    {
                        string dictKey = ranges.Cells[2, j].Text;
                        string dictValue = " ";
                        if (!(ranges.Cells[i, j].Text == null || ranges.Cells[i, j].Text == ""))
                        {
                            dictValue = ranges.Cells[i, j].Text;
                        }
                        //WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                        dict.Add(dictKey, dictValue);
                    }

                    Word.Document doc = null;

                    if (dict["num"].Substring(0, 1) == "3")
                    {
                        //合同制待遇一样
                        doc = doc1;
                    }
                    else
                    {
                        doc = doc1;
                    }

                    foreach (Word.Variable var in doc.Variables)
                    {
                        //WriteLine(var.Name);
                        var.Value = "";
                    }


                    if (!(dict["name"] == "" || dict["name"] == null))
                    {
                        WriteLine("OK");
                        foreach (String key in dict.Keys)
                        {
                            WriteLine($"KEY:{key}, Var:{dict[key]}");
                            doc.Variables.Add(key, dict[key]);
                        }
                        doc.Fields.Update();

                        doc.SaveAs2($@"{PathHeader}{dict["index"]}-{dict["name"]}-{dict["num"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                    }


                }
            }
            finally
            {
                doc1.Close();
                doc2.Close();
                thisWorkBook.Close(false);
                excelApp.Quit();
            }

        }
    }
}
