﻿using System;
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

namespace TuixiuVSTO.App
{
    class _ExcelToWord : IDisposable
    {
        #region 需要修改的变量
        public static string resultFileName = "解除劳动合同";
        public static string sumFileName = "data.xlsx";
        public static string templateFileName = "解除劳动合同.docx";
        #endregion

        #region 固定变量
        public static string workPath = @"D:\1\VSTO_ExcelToWord\";

        //起始位置, 注意修改！！！
        public static int startNum = 2;

        public int keyNum;

        #endregion

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document templateDocument;

        Excel.Range ranges;

        string PathHeader;

        public _ExcelToWord()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            templateDocument.Close(false);
            excelApp.Quit();
        }


        public void genForm(int rowNum = 0)
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

            openSheet2(workPath + sumFileName, out thisWorkBook, out thisWorkSheet, "Sheet1");

            try
            {
                ranges = thisWorkSheet.UsedRange;

                keyNum = ranges.Columns.Count;

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
                thisWorkBook.Close(false);
                excelApp.Quit();
                wordApp.Quit();
            }

        }

        private void openSheet2(string path, out Excel.Workbook workbook, out Excel.Worksheet worksheet, string sheetName)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("File cannot found");
                Application.Exit();
            }

            workbook = excelApp.Workbooks.Open(path);

            worksheet = workbook.Worksheets[sheetName];
        }

        private void openDoc2(string path, out Word.Document document)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("File cannot found");
                Application.Exit();
            }

            document = wordApp.Documents.Open(path);
        }

        private void doBatch(int rowNum)
        {
            openDoc2(workPath + templateFileName, out templateDocument);

            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= keyNum; j++)
            {
                string dictKey = ranges.Cells[startNum-1, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            if (!(dict["id"] == "" || dict["id"] == null))
            {
                /*
                string PathHeader = $@"{workPath}{dict["Name"]}({dict["Class"]})\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }
                */

                #region 填写表格

                foreach (string key in dict.Keys)
                {
                    templateDocument.Variables.Add(key, dict[key]);
                }

                //更新变量
                templateDocument.Fields.Update();

                //删除未填写的变量
                foreach (Field field in templateDocument.Fields)
                {
                    field.Select();
                    if (field.Result.Text == "错误!未提供文档变量。")
                    {
                        //WriteLine($"{field.Code.Text}");
                        field.Delete();
                    }
                }

                //重新更新变量
                templateDocument.Fields.Update();


                #endregion

                //templateDocument.SaveAs2($@"{PathHeader}【{string.Format("{0:d3}", dict["id"])}】{resultFileName}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                templateDocument.SaveAs2($@"{PathHeader}【{dict["name"]}】{resultFileName}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                templateDocument.Close(SaveChanges: false);

            }
        }
    }
}
