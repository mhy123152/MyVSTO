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

namespace TuixiuVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            Tuixiu tx = new Tuixiu();
            Gongzi gz = new Gongzi();
            //tx.mergeExcelSheet();

            gz.genSheets();
        }

    }

    class Tuixiu
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

        public void tuixiu()
        {
            string workPath = @"D:\Library\Desktop\1\tuixiu\";

            string Path0 = workPath + "t0.docx";
            string Path1 = workPath + "t1.docx";
            string Path2 = workPath + "t2.docx";
            string Path3 = workPath + "t3.docx";
            string Path4 = workPath + "t4.docx";
            string Path5 = workPath + "t5.docx";

            //模板数量
            int templateNum = 6;

            //字段数量
            int keyNum = 11;

            //起始位置, 注意修改！！！
            int startNum = 3;

            string sumFileName = "2018年退休.xlsm";

            string[] Paths = new string[templateNum];
            for (int i = 0; i < Paths.Length; i++)
            {
                Paths[i] = $@"{workPath}t{i}.docx";
            }

            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();


            if (!File.Exists(workPath + sumFileName))
            {
                MessageBox.Show("File cannot found");
                return;
            }

            Excel.Workbook thisWorkBook = null;
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


            Excel.Worksheet thisWorkSheet = null;
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


            Word.Document[] docs = new Word.Document[templateNum];
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
                Excel.Range ranges = thisWorkSheet.UsedRange;

                string PathHeader = $@"{workPath}result\";
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

                for (int i = startNum; i <= ranges.Rows.Count; i++)
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
                        string dictValue = ranges.Cells[i, j].Text;
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

        public void hetong()
        {

            string workPath = @"D:\Library\Desktop\1\hetong\";
            string sumFileName = "Data.xlsx";
            string templateFileName = "事业单位聘用合同（人事代理续签模板）.docx";
            string templateFileName2 = "事业单位聘用合同（合同制续签模板）.docx";
            int KeyNum = 14;

            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();


            if (!File.Exists(workPath + templateFileName))
            {
                MessageBox.Show("File cannot found");
                return;
            }

            Excel.Workbook thisWorkBook = null;
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


            Excel.Worksheet thisWorkSheet = null;
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

                Excel.Range ranges = thisWorkSheet.UsedRange;

                string PathHeader = $@"{workPath}result\";
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

                        doc.SaveAs2($@"{PathHeader}{dict["num"]}{dict["name"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

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

        public void zhenggao2018()
        {

            string workPath = @"D:\Library\Desktop\1\2018zhenggao\";
            string sumFileName = "2018年正高女职工退休.xlsm";
            string templateFileName = "template3.docx";
            int KeyNum = 4;


            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();


            if (!File.Exists(workPath + templateFileName))
            {
                MessageBox.Show("File cannot found");
                return;
            }

            Excel.Workbook thisWorkBook = null;
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


            Excel.Worksheet thisWorkSheet = null;
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


            Word.Document doc = null;
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

                Excel.Range ranges = thisWorkSheet.UsedRange;

                /*
                string Name = "";
                string Year = "";
                string Month = "";
                */

                string PathHeader = $@"{workPath}result\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }

                for (int i = 3; i <= ranges.Rows.Count; i++)
                {
                    foreach (Word.Variable var in doc.Variables)
                    {
                        WriteLine(var.Name);
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
                        foreach (String key in dict.Keys)
                        {
                            WriteLine($"KEY:{key}, Var:{dict[key]}");
                            doc.Variables.Add(key, dict[key]);
                        }
                        doc.Fields.Update();

                        doc.SaveAs2($@"{PathHeader}{dict["Year"]}{dict["Month"]}{dict["Name"]}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

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

    class Gongzi
    {
        public void genSheets()
        {
            int incRow = 9;
            int itemPerPage = 3;
            int curRow = 1;
            string workPath = @"D:\1\";
            string sumFileName = "template.xlsx";

            Excel.Application excelApp = new Excel.Application();
            //Excel.Workbook thisWorkBook = excelApp.ActiveWorkbook;
            //Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(excelApp.ActiveWorkbook.Path + "\\a.xlsx");
            Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            Excel.Worksheet thisWorkSheet = thisWorkBook.Worksheets["Sheet1"];

            //workPath = thisWorkBook.Path;
            WriteLine(workPath);


            using (var db = new DBModel())
            {
                var c2017 =
                    db.C2017sum
                    .OrderBy(a => a.机构编号)
                    .ThenByDescending(a => a.年合计收入)
                    .ToList();

                int itemCount = c2017.Count;
                //int itemCount = 11;

                foreach (var item in c2017)
                //for (int i = 0; i < itemCount; i++)
                {
                    //var item = c2017[i];
                    WriteLine($"{item.机构}, {item.姓名}, {item.职工号}, {item.身份证号码}, {item.年合计收入}");
                    curRow = toSheet2(item, thisWorkSheet, curRow, incRow, itemCount, itemPerPage);
                }

            }

            thisWorkBook.SaveAs(Filename: $@"{workPath}done.xlsx");
            thisWorkBook.Close(false);
        }

        public int toSheet2(C2017sum item, Excel.Worksheet thisWorkSheet, int curRow, int incRow, int itemCount, int itemPerPage)
        {
            thisWorkSheet.Cells[2, 2] = item.机构;
            thisWorkSheet.Cells[3, 2] = item.姓名;
            thisWorkSheet.Cells[3, 4] = item.职工号;
            thisWorkSheet.Cells[4, 2] = item.身份证号码;
            thisWorkSheet.Cells[7, 1] = item.年合计收入;
            thisWorkSheet.Cells[7, 2] = item.工资;
            thisWorkSheet.Cells[7, 3] = item.奖金;
            thisWorkSheet.Cells[7, 4] = item.奖励性补贴;

            string srcStr = "1:9";
            var srcSheet = thisWorkSheet.Range[srcStr];

            curRow = curRow + incRow * itemPerPage;

            if (curRow >= (itemCount + 1) * incRow + 1)
            {
                int pageNum = curRow / (incRow * itemPerPage);
                WriteLine($"分页：{pageNum}");
                curRow = curRow - pageNum * itemPerPage * incRow + incRow;
            }

            string dstStr = $"{curRow}:{curRow + incRow - 1}";
            WriteLine($"src:{srcStr}, dst:{dstStr}");

            var dstSheet = thisWorkSheet.Range[dstStr];

            srcSheet.Copy(dstSheet);

            return curRow;
        }

        public int toSheet(C2017sum item, Excel.Worksheet thisWorkSheet, int curRow, int incRow)
        {
            thisWorkSheet.Cells[2, 2] = item.机构;
            thisWorkSheet.Cells[3, 2] = item.姓名;
            thisWorkSheet.Cells[3, 4] = item.职工号;
            thisWorkSheet.Cells[4, 2] = item.身份证号码;
            thisWorkSheet.Cells[7, 1] = item.年合计收入;
            thisWorkSheet.Cells[7, 2] = item.工资;
            thisWorkSheet.Cells[7, 3] = item.奖金;
            thisWorkSheet.Cells[7, 4] = item.奖励性补贴;

            string srcStr = "1:9";
            var srcSheet = thisWorkSheet.Range[srcStr];

            curRow = curRow + incRow;

            string dstStr = $"{curRow}:{curRow + incRow - 1}";
            var dstSheet = thisWorkSheet.Range[dstStr];

            WriteLine($"src:{srcStr}, dst:{dstStr}");

            srcSheet.Copy(dstSheet);

            return curRow;
        }
    }
}
