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
    class zgzm : IDisposable
    {
        public static string workPath = @"D:\1\职工证明\";

        //模板数量
        public static int templateNum = 1;

        //字段数量
        public static int keyNum = 3;

        //起始位置, 注意修改！！！
        public static int startNum = 3;

        //退休人员信息的文件名
        public static string sumFileName = "data.xlsx";

        string[] Paths = new string[templateNum];

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document[] docs;
        Excel.Range ranges;
        string PathHeader;

        public zgzm()
        {
            for (int i = 0; i < Paths.Length; i++)
            {
                Paths[i] = $@"{workPath}t{i}.docx";
            }

            excelApp = new Excel.Application();
            wordApp = new Word.Application();
            docs = new Word.Document[templateNum];
        }

        public void Dispose()
        {
            excelApp.Quit();
            wordApp.Quit();
            docs = null;
        }


        public void run(int rowNum = 0)
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


            for (int i = 0; i < docs.Length; i++)
            {
                try
                {
                    docs[i] = wordApp.Documents.Open(Paths[i]);
                }
                catch
                {
                    if (docs[i] != null)
                    {
                        docs[i].Close();
                    }
                    MessageBox.Show("Word Path not found");
                    return;
                }

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

                /*
                string Name = "";
                string RetireDate = "";
                string RetireDate2 = "";
                string MyDate = "";
                string Class = "";
                string WorkDate = "";
                string Level = "";
                string Title = "";
                string Chief = "";
                string Cadre = "";
                */

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
                foreach (Word.Document doc in docs)
                {
                    doc.Close();
                }
                thisWorkBook.Close(false);
                excelApp.Quit();
                wordApp.Quit();
            }

        }

        private void doBatch(int rowNum)
        {
            foreach (Word.Document doc in docs)
            {
                foreach (Word.Variable var in doc.Variables)
                {
                    var.Delete();
                }
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= keyNum; j++)
            {
                string dictKey = ranges.Cells[2, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            /*
            if(dict["isRun"] != "a")
            {
                Dispose();
                //Environment.Exit(0);
            }
            */

            if (!(dict["Name"] == "" || dict["Name"] == null))
            {
                /*
                string PathHeader = $@"{workPath}{dict["Name"]}({dict["Class"]})\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }
                */

                docs[0].Variables.Add("Name", dict["Name"]);
                docs[0].Variables.Add("Sex", dict["Sex"]);
                docs[0].Variables.Add("uID", dict["uID"]);
                //docs[0].Variables.Add("Date", ConvertDateTimeToDate(DateTime.Now.ToLongDateString(), "zh-CN"));

                foreach (Word.Document doc in docs)
                {
                    /*
                    foreach (Word.Variable var in doc.Variables)
                    {
                        var.Value = dict[var.Name];
                    }
                    */

                    doc.Fields.Update();
                }

                docs[0].SaveAs2($@"{PathHeader}{dict["Name"]}-职工证明.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

 
            }
        }

        public static string ConvertDateTimeToDate(string dateTimeString, String langCulture)
        {

            CultureInfo culture = new CultureInfo(langCulture);
            DateTime dt = DateTime.MinValue;

            if (DateTime.TryParse(dateTimeString, out dt))
            {
                return dt.ToString("D", culture);
            }
            return dateTimeString;
        }


    }
}
