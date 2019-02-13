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
    class Tuixiu : IDisposable
    {
        public static string workPath = @"D:\Library\Desktop\1\tuixiu\";
        public static string Path0 = workPath + "t0.docx";
        public static string Path1 = workPath + "t1.docx";
        public static string Path2 = workPath + "t2.docx";
        public static string Path3 = workPath + "t3.docx";
        public static string Path4 = workPath + "t4.docx";
        public static string Path5 = workPath + "t5.docx";

        //模板数量
        public static int templateNum = 6;

        //字段数量
        public static int keyNum = 11;

        //起始位置, 注意修改！！！
        public static int startNum = 3;

        //退休人员信息的文件名
        public static string sumFileName = "2019年退休.xlsm";

        string[] Paths = new string[templateNum];

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document[] docs;
        Excel.Range ranges;
        string PathHeader;

        public Tuixiu()
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

            for (int j = 1; j < keyNum; j++)
            {
                string dictKey = ranges.Cells[2, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

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
                docs[0].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[0].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[0].Variables.Add("Date", dict["MyDate"]);

                docs[1].Variables.Add("Name", dict["Name"]);
                docs[1].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[1].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[1].Variables.Add("Date", dict["MyDate"]);

                docs[2].Variables.Add("Class", dict["Class"]);
                docs[2].Variables.Add("Name", dict["Name"]);
                docs[2].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[2].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[2].Variables.Add("Date", dict["MyDate"]);

                docs[3].Variables.Add("Name", dict["Name"]);
                docs[3].Variables.Add("Class", dict["Class"]);
                docs[3].Variables.Add("WorkDate", dict["WorkDate"]);
                docs[3].Variables.Add("Level", dict["Level"]);
                docs[3].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[3].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[3].Variables.Add("Date", dict["MyDate"]);

                docs[4].Variables.Add("Name", dict["Name"]);
                docs[4].Variables.Add("Class", dict["Class"]);
                docs[4].Variables.Add("WorkDate", dict["WorkDate"]);
                docs[4].Variables.Add("Level", dict["Level"]);
                docs[4].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[4].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[4].Variables.Add("Date", dict["MyDate"]);

                docs[5].Variables.Add("Title", dict["Title"]);
                docs[5].Variables.Add("Name", dict["Name"]);
                docs[5].Variables.Add("Class", dict["Class"]);
                docs[5].Variables.Add("WorkDate", dict["WorkDate"]);
                docs[5].Variables.Add("Level", dict["Level"]);
                docs[5].Variables.Add("RetireDate", dict["RetireDate"]);
                docs[5].Variables.Add("RetireDate2", dict["RetireDate2"]);
                docs[5].Variables.Add("Date", dict["MyDate"]);

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

                //如果是高级职称则用模板t0否则用模板t1
                if (dict["Chief"] == "1")
                {
                    docs[0].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}1_个人.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                }
                else
                {
                    docs[1].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}1_个人.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                }

                //判断是否是退养人员
                if (!(dict["Class"] == "" || dict["Class"] == null || dict["Class"] == "退养"))
                {
                    docs[2].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}2_科室.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                }

                docs[3].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}3_工会.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                docs[4].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}4_老干.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                //判断是否需要报医务、护理或院办
                if (!(dict["Title"] == "" || dict["Title"] == null || dict["Title"] == "无"))
                {
                    docs[5].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}5_{dict["Title"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                }

                //如果是干部，则再次把模板t5的Title改为党办，然后另存
                if (dict["Cadre"] == "1")
                {
                    docs[5].Variables["Title"].Value = "党办";
                    docs[5].Fields.Update();
                    docs[5].SaveAs2($@"{PathHeader}{dict["RetireDate"]}-{dict["Name"]}6_党办.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                }
            }
        }


        


        public void test2()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook currentWorkBook = excelApp.ActiveWorkbook;

            string workPath = currentWorkBook.Path;
            WriteLine(workPath);

            string aPath = workPath + "\\a.xlsx";
            WriteLine(aPath);

            Excel.Workbook workbook = excelApp.Workbooks.Open(aPath);

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                WriteLine(sheet.Name);
                Excel.Range range = sheet.Range["A1"];
                string value = range.Value2;
                WriteLine(value);

                Excel.Range range2 = sheet.Range["A1", "B3"];
                object[,] values = range2.Value2;
                string v22 = values.GetValue(2, 2) as string;
                WriteLine(v22);
            }
        }

        public void test()
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); //获取打开了的Excel应用程序

        }

        public void mergeExcelSheet()
        {
            string workPath = @"D:\Library\Desktop\1\merge\";
            string sumFileName = "sum.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            Excel.Worksheet thisWorkSheet = thisWorkBook.Worksheets["Sheet1"];

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

                    Excel.Worksheet copyWorkSheet = workbook.Worksheets["Sheet1"];

                    if (copyWorkSheet != null)
                    {
                        copyWorkSheet.Name = Path.GetFileNameWithoutExtension(file.FullName);
                        copyWorkSheet.Copy(thisWorkSheet);

                    }
                    else
                    {
                        WriteLine($"Sheet1 Not Found: ${file.Name}");
                    }

                    workbook.Close(false);

                }
            }

            thisWorkBook.SaveAs(Filename: $@"{workPath}done.xlsx");
            thisWorkBook.Close();
        }


    }
}
