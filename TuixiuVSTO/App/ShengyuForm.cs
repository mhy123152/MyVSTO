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
    class ShengyuForm : IDisposable
    {
        public static string workPath = @"D:\1\shengyuform\";

        //起始位置, 注意修改！！！
        public static int startNum = 2;

        //退休人员信息的文件名
        public static string sumFileName = "生育保险待遇申领表数据.xlsx";
        public static string templateFileName = "生育保险待遇申领表模板.xlsx";

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook templateWorkBook;
        Excel.Worksheet templateWorkSheet;
        Excel.Worksheet templateWorkSheet2;

        Excel.Range ranges;
        string PathHeader;

        public ShengyuForm()
        {
            excelApp = new Excel.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            //templateWorkBook.Close(false);

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            excelApp = null;
        }

        public void genForm()
        {
            genForm(new int[] { 0 });
        }

        public void genForm(int[] rowNums)
        {
            PathHeader = $@"{workPath}result\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            Directory.CreateDirectory(workPath);
            Directory.CreateDirectory(PathHeader);


            openSheet2(workPath + sumFileName, out thisWorkBook, out thisWorkSheet, "Sheet1");

            try
            {
                ranges = thisWorkSheet.UsedRange;

                WriteLine(PathHeader);
                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }

                if (rowNums[0] == 0)
                {
                    for (int i = startNum; i <= ranges.Rows.Count; i++)
                    {
                        doBatch(i);
                    }
                }
                else
                {
                    foreach (int rowNum in rowNums)
                    {
                        doBatch(rowNum);
                    }
                }

            }
            finally
            {
                Dispose();
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

        private void doBatch(int rowNum)
        {

            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= ranges.Columns.Count; j++)
            {
                string dictKey = ranges.Cells[1, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }


            if (!(dict["姓名"] == "" || dict["姓名"] == null))
            {
                openSheet2(workPath + templateFileName, out templateWorkBook, out templateWorkSheet, "Sheet1");

                #region 填写表格

                templateWorkSheet.Range["B3"].Value2 = dict["姓名"];
                templateWorkSheet.Range["H3"].Value2 = dict["身份证号"];
                templateWorkSheet.Range["B5"].Value2 = dict["入院日期"];
                templateWorkSheet.Range["D5"].Value2 = dict["出院日期"];
                templateWorkSheet.Range["G7"].Value2 = dict["手机号"];
                templateWorkSheet.Range["B6"].Value2 = dict["住院费用"];
                templateWorkSheet.Range["D4"].Value2 = dict["生育时间"];

                #endregion

                templateWorkBook.SaveAs(Filename: $@"{PathHeader}{dict["序号"]}_{dict["姓名"]}_{dict["身份证号"]}.xlsx");

                templateWorkBook.Close(SaveChanges: false);

            }
        }

    }
}
