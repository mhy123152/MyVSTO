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
using System.Globalization;

namespace TuixiuVSTO.App
{
    class TuixiuFormNew : IDisposable
    {
        public static string workPath = @"D:\1\tuixiuformNew\";

        //字段数量
        public static int keyNum = 30;

        //起始位置, 注意修改！！！
        public static int startNum = 2;

        //退休人员信息的文件名
        public static string sumFileName = "退休审批表数据.xlsx";
        public static string templateFileName = "退休审批表（2019事业）.xlsx";

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook templateWorkBook;
        Excel.Worksheet templateWorkSheet;
        Excel.Worksheet templateWorkSheet2;

        Excel.Range ranges;
        string PathHeader;
        string PathHeader1;
        string PathHeader2;

        public TuixiuFormNew()
        {
            excelApp = new Excel.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            templateWorkBook.Close(false);
            excelApp.Quit();
        }

        DateTime dtMax = new DateTime(2999, 1, 1);
        DateTime dtMin = new DateTime(1900, 1, 1);

        public void genForm()
        {
            genForm(dtMax, dtMin, new int[] { 0 });
        }

        public void genForm(int[] rowNums)
        {
            genForm(dtMax, dtMin, rowNums);
        }

        public void genForm(DateTime dateBefore, DateTime dateAfter)
        {
            genForm(dateBefore, dateAfter, new int[] { 0 });
        }

        public void genForm(DateTime dateBefore, DateTime dateAfter, int[] rowNums)
        {
            PathHeader = $@"{workPath}result\";
            PathHeader1 = $@"{workPath}result\有独生子女证\";
            PathHeader2 = $@"{workPath}result\有婚育情况证明书\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            Directory.CreateDirectory(workPath)
                .Refresh();

            Directory.CreateDirectory(PathHeader)
                .Refresh();
            //Directory.CreateDirectory(PathHeader1);
            //Directory.CreateDirectory(PathHeader2);


            openSheet2(workPath + sumFileName, out thisWorkBook, out thisWorkSheet, "Sheet1");

            try
            {
                ranges = thisWorkSheet.UsedRange;

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

                if (rowNums[0] == 0)
                {
                    for (int i = startNum; i <= ranges.Rows.Count; i++)
                    {
                        doBatch(i, dateBefore, dateAfter);
                    }
                }
                else
                {
                    foreach(int rowNum in rowNums)
                    {
                        doBatch(rowNum, dtMax, dtMin);
                    }
                }

            }
            finally
            {
                thisWorkBook.Close(false);
                //templateWorkBook.Close(false);
                excelApp.Quit();
            }

        }

        //private void openSheet(string path, out Excel.Workbook workbook, out Excel.Worksheet worksheet)
        //{
        //    openSheet(path, out workbook, out worksheet, "Sheet1");
        //}

        //private void openSheet(string path, out Excel.Workbook workbook, out Excel.Worksheet worksheet, string sheetName)
        //{
        //    if (!File.Exists(path))
        //    {
        //        MessageBox.Show("File cannot found");
        //        Application.Exit();
        //    }

        //    try
        //    {
        //        workbook = excelApp.Workbooks.Open(path);
        //    }
        //    catch
        //    {
        //        excelApp.Quit();
        //        MessageBox.Show("Excel Path not found");
        //        Application.Exit();
        //    }

        //    try
        //    {
        //        worksheet = workbook.Worksheets[sheetName];
        //    }
        //    catch
        //    {
        //        workbook.Close(false);
        //        excelApp.Quit();
        //        MessageBox.Show("Sheet not found");
        //        Application.Exit();
        //    }
        //}

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

        private void doBatch(int rowNum, DateTime dateBefore, DateTime dateAfter)
        {

            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= keyNum; j++)
            {
                string dictKey = ranges.Cells[1, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo
            {
                ShortDatePattern = "yyyy/MM/dd"
            };
            DateTime dt = Convert.ToDateTime(dict["退休时间"], dtFormat);

            if (!(dateBefore >= dt && dateAfter <= dt))
            {
                WriteLine($"{dict["姓名"]} PASS");
                return;
            }


            if (!(dict["姓名"] == "" || dict["姓名"] == null))
            {
                openSheet2(workPath + templateFileName, out templateWorkBook, out templateWorkSheet, "审批表正面");

                templateWorkSheet2 = templateWorkBook.Worksheets["审批表背面"];

                /*
                string PathHeader = $@"{workPath}{dict["Name"]}({dict["Class"]})\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }
                */

                #region 填写表格

                templateWorkSheet.Range["D2"].Value2 = dict["个人编号"];
                templateWorkSheet.Range["D3"].Value2 = dict["姓名"];
                templateWorkSheet.Range["K3"].Value2 = dict["性别"];
                templateWorkSheet.Range["K4"].Value2 = dict["2014年9月岗位"];
                templateWorkSheet.Range["D5"].Value2 = dict["出生时间"];
                templateWorkSheet.Range["K5"].Value2 = dict["参加工作时间"];
                templateWorkSheet.Range["O5"].Value2 = dict["退休时间"];
                templateWorkSheet.Range["D6"].Value2 = "荆州市中心医院";
                templateWorkSheet.Range["O6"].Value2 = dict["工作年限"];
                templateWorkSheet.Range["D7"].Value2 = dict["身份证号码"];
                templateWorkSheet.Range["F9"].Value2 = dict["2014年9月薪级"];
                templateWorkSheet.Range["F11"].Value2 = dict["2014年9月护教10%"];
                templateWorkSheet.Range["G12"].Value2 = dict["退休时岗位"];
                templateWorkSheet.Range["G13"].Value2 = dict["退休时薪级"];
                templateWorkSheet.Range["H15"].Value2 = "5";

                #endregion

                string fn;

                fn = $@"【{dict["退休时间"].Substring(0, 7)}】【{dict["档案编号"]}】{dict["姓名"]}";

                //if (dict["是否女干"] == "是")
                //{
                //    fn = $@"【女干】【{dict["退休时间"].Substring(0, 7)}】【{dict["档案编号"]}】{dict["姓名"]}";
                //}
                //else
                //{
                //    fn = $@"【正常】【{dict["退休时间"].Substring(0, 7)}】【{dict["档案编号"]}】{dict["姓名"]}";
                //}

                //string picPath = $@"{workPath}独生子女证\{fn}.png";

                //if (File.Exists(picPath))
                //{
                //    templateWorkSheet2.Shapes.AddPicture(picPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 450, -1, -1);

                //    //templateWorkBook.SaveAs(Filename: $@"{PathHeader1}{fn}.xlsx");
                //}
                //else if (File.Exists(picPath2))
                //{
                //    templateWorkSheet2.Shapes.AddPicture(picPath2, MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 450, -1, -1);

                //    //templateWorkBook.SaveAs(Filename: $@"{PathHeader1}{fn}.xlsx");
                //}
                //else if (File.Exists($@"{workPath}婚育情况证明书\{fn}.png"))
                //{
                //    //templateWorkBook.SaveAs(Filename: $@"{PathHeader2}{fn}.xlsx");
                //}
                ////else
                ////{
                ////    templateWorkBook.SaveAs(Filename: $@"{PathHeader}{fn}.xlsx");
                ////}

                templateWorkBook.SaveAs(Filename: $@"{PathHeader}{fn}.xlsx");


                templateWorkBook.Close(SaveChanges: false);

            }
        }

    }
}
