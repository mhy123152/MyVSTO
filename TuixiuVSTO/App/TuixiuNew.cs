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
    class TuixiuNew : IDisposable
    {
        public static string workPath = @"C:\Users\Michael\OneDrive\0Working\退休和辞职\tuixiu\";
        public static string templatePath = workPath + "退休通知模板.docx";

        //起始位置, 注意修改！！！
        public static int startNum = 3;

        //退休人员信息的文件名
        public static string sumFileName = "2022年退休.xlsx";

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document templateDocument;

        Excel.Range ranges;
        string PathHeader;

        public TuixiuNew()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            excelApp.Quit();
            wordApp.Quit();
        }


        public void tuixiu(int rowNum = 0)
        {

            if (!File.Exists(workPath + sumFileName))
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
                templateDocument = wordApp.Documents.Open(templatePath);
            }
            catch
            {
                if (templateDocument != null)
                {
                    templateDocument.Close();
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

                if (rowNum == 0)
                {
                    for (int i = startNum; i <= ranges.Rows.Count; i++)
                    {
                        doBatch(i);
                    }
                }
                else
                {
                    doBatch(rowNum);
                }

            }
            finally
            {

                templateDocument.Close();

                thisWorkBook.Close(false);
                excelApp.Quit();
                wordApp.Quit();
            }

        }

        private void doBatch(int rowNum)
        {

            foreach (Word.Variable var in templateDocument.Variables)
            {
                var.Delete();
            }


            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= ranges.Columns.Count; j++)
            {
                string dictKey = ranges.Cells[2, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            if (!(dict["Name"] == "" || dict["Name"] == null))
            {

                templateDocument.Variables.Add("Name", dict["Name"]);
                templateDocument.Variables.Add("BirthDate", dict["BirthDate"]);
                templateDocument.Variables.Add("WorkDate", dict["WorkDate"]);
                templateDocument.Variables.Add("Level", dict["Level"]);
                templateDocument.Variables.Add("RetireDate", dict["RetireDate"]);
                templateDocument.Variables.Add("RetireDate2", dict["RetireDate2"]);
                templateDocument.Variables.Add("MyDate", dict["MyDate"]);


                templateDocument.Fields.Update();

                templateDocument.SaveAs2($@"{PathHeader}退休通知_{dict["RetireDate"]}_{dict["Name"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

            }
        }

    }
}
