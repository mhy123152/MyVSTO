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

namespace TuixiuVSTO.App
{
    class BianziADD : IDisposable
    {

        public static string workPath = @"D:\1\新增编制\";

        //字段数量, 注意修改！！！
        public static int keyNum = 9;

        //起始位置, 注意修改！！！
        public static int startNum = 3;

        //退休人员信息的文件名
        public static string sumFileName = "新增编制数据.xlsx";
        public static string templateFileName = "新增编制招聘模板.docx";

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document templateDocument;

        Excel.Range ranges;

        string PathHeader;

        public BianziADD()
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
                string dictKey = ranges.Cells[2, j].Text;
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
                    /*
                    不填写以下数据：
                    class_zj;class_gq, 根据情况打钩或不打钩
                    */
                    if (key == "class_zj" || key == "class_gq")
                    {
                        //Pass
                    }
                    else
                    {
                        templateDocument.Variables.Add(key, dict[key]);
                    }
                }

                if (dict["class"] == "专技")
                {
                    templateDocument.Variables.Add("class_zj", "☑");
                    templateDocument.Variables.Add("class_gq", "□");
                }
                else if(dict["class"] == "工勤")
                {
                    templateDocument.Variables.Add("class_gq", "☑");
                    templateDocument.Variables.Add("class_zj", "□");
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

                templateDocument.SaveAs2($@"{PathHeader}【{dict["id"]}】【{dict["name"]}】招聘登记表.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                templateDocument.Close(SaveChanges: false);

            }
        }
    }
}
