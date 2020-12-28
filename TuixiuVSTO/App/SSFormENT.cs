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
using Syncfusion.XlsIO;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using System.Text.RegularExpressions;

namespace TuixiuVSTO.App
{
    class SSFormENT : IDisposable
    {
        //public static string workPath = @"D:\1\ssform500\";
        public static string workPath = @"D:\1\SSFormENT\";

        //执行年月开始行号
        public static int dataStartNum = 14;

        //执行年月结束行号
        public static int dataEndNum = 33;

        //每页行数, 注意修改！！！
        public static int pageRowCount = 33;

        //汇总表格的标题行号
        private int sumFileCurRow = 1;

        //编号
        private int id = 0;

        //退休人员信息的文件名
        public static string dataFileName = "荆州市中心医院(编外人员工资)在职沿革表.xlsx";
        public static string sumFileName = "template.xlsx";
        public static string sssumFileName = "编外人员历史沿革表汇总数据.xlsx";

        public static string pdfFileName = "荆州市中心医院(岗位设置)退休人员沿革表.pdf";
        public static string resultPdfFileName = "中人历史沿革表(已排序).pdf";

        public string[] names;
        public string[] levels;
        public int[] levelsPay;
        public int[] wuyePay;

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook sumWorkBook;
        Excel.Worksheet sumWorkSheet;

        Excel.Workbook sumWorkBook2;
        //Excel.Worksheet sumWorkSheet2;

        Excel.Range ranges;
        Excel.Range searchRanges;

        DateTimeFormatInfo dtFormat;

        string previousName;

        public SSFormENT()
        {
            excelApp = new Excel.Application();

            dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";

            levels = new[] { "正高级", "副高级", "中级", "助理级", "员级", "见习期、初期", "高级技师", "技师", "高级工", "中级工", "初级工" };

            levelsPay = new[] { 1134, 952, 872, 891, 822, 822, 943, 881, 817, 846, 813 };
            wuyePay = new[] { 240, 240, 200, 200, 200, 200, 200, 200, 160, 160, 160 };

            names = new[] { "" };

        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            sumWorkBook.Close(false);
            excelApp.Quit();
        }


        public void genForm(int pageNum = 0)
        {
            string sumFilePath = $@"{workPath}{sssumFileName}";

            if (File.Exists(sumFilePath))
            {
                File.Copy(sumFilePath, $@"{workPath}old{sssumFileName}", true);
                File.Delete(sumFilePath);
            }


            openSheet2(workPath + dataFileName, out thisWorkBook, out thisWorkSheet, "事业在职沿革表");

            openSheet2(workPath + sumFileName, out sumWorkBook, out sumWorkSheet, "Sheet1");

            try
            {
                ranges = thisWorkSheet.UsedRange;


                if (pageNum == 0)
                {
                    for (int page = 1; page <= ranges.Rows.Count / pageRowCount; page++)
                    {
                        doBatch(page);
                    }
                }
                else
                {
                    doBatch(pageNum);
                }


                sumWorkBook.SaveAs(Filename: sumFilePath);

            }
            finally
            {
                thisWorkBook.Close(false);
                sumWorkBook.Close(false);
                excelApp.Quit();
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

        private void doBatch(int pageNum)
        {
            int firstRow = (pageNum - 1) * pageRowCount + 1;
            int myDataStartNum = firstRow + dataStartNum - 1;
            int myDataEndNum = firstRow + dataEndNum - 1;

            string name = ranges.Cells[firstRow + 2, 2].Text;

            if (!(name == "" || name == null))
            {
                if (previousName != name)
                {
                    id += 1;

                    WriteLine($"{id}:{name}");

                    DateTime previousExecuteDateTime = Convert.ToDateTime("1990-01-01", dtFormat);

                    for (int i = myDataStartNum; i <= myDataEndNum; i++)
                    {


                        string executeDateCellValue = ranges.Cells[i, 2].Text;

                        if (!(executeDateCellValue == "" || executeDateCellValue == null))
                        {
                            string levelCellValue = ranges.Cells[i, 3].Text;

                            int levelPay = 0;
                            int wyPay = 0;

                            for (int j = 0; j < levels.Length; j++)
                            {
                                if (levelCellValue.Contains($"{levels[j]}("))
                                {
                                    levelPay = levelsPay[j];
                                    wyPay = wuyePay[j];
                                }
                            }

                            string executeDate = $"{executeDateCellValue.Replace(@".", "-")}-01";

                            DateTime executeDateTime = Convert.ToDateTime(executeDate, dtFormat);

                            if (executeDateTime.Year != previousExecuteDateTime.Year)
                            {
                                sumFileCurRow += 1;
                            }

                            #region 填写表格

                            sumWorkSheet.Cells[sumFileCurRow, 1].Value2 = id;                       //编号
                            sumWorkSheet.Cells[sumFileCurRow, 2].Value2 = name;                     //姓名
                            sumWorkSheet.Cells[sumFileCurRow, 3].Value2 = executeDateTime;          //执行日期
                            sumWorkSheet.Cells[sumFileCurRow, 4].Value2 = ranges.Cells[i, 3].Text;  //岗位
                            sumWorkSheet.Cells[sumFileCurRow, 5].Value2 = ranges.Cells[i, 4].Text;  //薪级
                            sumWorkSheet.Cells[sumFileCurRow, 6].Value2 = ranges.Cells[i, 5].Text;  //岗位工资
                            sumWorkSheet.Cells[sumFileCurRow, 7].Value2 = ranges.Cells[i, 6].Text;  //薪级工资
                            sumWorkSheet.Cells[sumFileCurRow, 8].Value2 = levelPay;                 //基础绩效工资
                            sumWorkSheet.Cells[sumFileCurRow, 9].Value2 = ranges.Cells[i, 7].Text;  //教护10%
                            sumWorkSheet.Cells[sumFileCurRow, 10].Value2 = 50;  //住房补贴
                            sumWorkSheet.Cells[sumFileCurRow, 11].Value2 = wyPay;  //物业补贴
                            sumWorkSheet.Cells[sumFileCurRow, 12].Value2 = 56;  //保留津补贴

                            #endregion


                            previousExecuteDateTime = executeDateTime;
                        }
                    }


                }

            }

            previousName = name;

        }

    }
}

