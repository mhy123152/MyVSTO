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
    class SSDataFromSumForm : IDisposable
    {
        public static string workPath = @"D:\1\SSDataFromSumForm\";

        //汇总表格的标题行号
        private int eduCurRow = 2;
        private int levelCurRow = 2;

        //退休人员信息的文件名
        public static string dataFileName = "非在编人员信息汇总表.xlsx";
        public static string sumFileName = "帅手数据模板.xlsx";
        public static string sssumFileName = "帅手数据.xlsx";

        public static string pdfFileName = "荆州市中心医院(岗位设置)退休人员沿革表.pdf";
        public static string resultPdfFileName = "中人历史沿革表(已排序).pdf";

        //起始位置, 注意修改！！！
        public static int startNum = 2;

        public int keyNum;

        public string[] names;
        public string[] levels;
        public string[] eduList;
        public string[] ZJlevelList;
        public string[] GQlevelList;

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook sumWorkBook;
        Excel.Worksheet sumWorkSheet;

        Excel.Workbook sumWorkBook2;
        Excel.Worksheet sumWorkSheet2;

        Excel.Range ranges;
        Excel.Range searchRanges;

        DateTimeFormatInfo dtFormat;

        string previousName;

        public SSDataFromSumForm()
        {
            excelApp = new Excel.Application();

            dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";

            levels = new[] { "正高级", "副高级", "中级", "助理级", "员级", "见习期、初期", "高级技师", "技师", "高级工", "中级工", "初级工" };

            names = new[] { "" };

            eduList = new[] { "中专", "大学专科", "大学本科", "硕士研究生", "博士研究生" };
            ZJlevelList = new[] { "员级", "助理级", "中级", "副高级", "正高级" };
            GQlevelList = new[] { "初级工", "中级工", "高级工", "技师", "高级技师" };

        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            sumWorkBook.Close(false);
            excelApp.Quit();
        }


        public void genForm(int rowNum = 0)
        {
            string sumFilePath = $@"{workPath}{sssumFileName}";

            if (File.Exists(sumFilePath))
            {
                File.Copy(sumFilePath, $@"{workPath}old{sssumFileName}", true);
                File.Delete(sumFilePath);
            }


            openSheet2(workPath + dataFileName, out thisWorkBook, out thisWorkSheet, "Sheet1");

            openSheet2(workPath + sumFileName, out sumWorkBook, out sumWorkSheet, "学历表");
            openSheet2(workPath + sumFileName, out sumWorkBook2, out sumWorkSheet2, "简历表");

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

        private void doBatch(int rowNum)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= keyNum; j++)
            {
                string dictKey = ranges.Cells[startNum - 1, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                //WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            WriteLine($"{dict["序号"]}\t{dict["姓名"]}  \t{dict["出生日期"]}\t{dict["参加工作时间"]}\t{dict["职务类别"]}\t{dict["是否护士"]}");

            string isNurse = "";
            if (dict["是否护士"] == "是")
            {
                isNurse = "02;";
            }


            foreach (string edu in eduList)
            {
                if (dict[edu] != null && dict[edu] != "")
                {
                    eduCurRow += 1;

                    #region 填写学历表

                    sumWorkSheet.Cells[eduCurRow, 1].Value2 = dict["序号"];
                    sumWorkSheet.Cells[eduCurRow, 2].Value2 = dict[edu];              //取得学历时间
                    sumWorkSheet.Cells[eduCurRow, 3].Value2 = edu;                    //学历
                    sumWorkSheet.Cells[eduCurRow, 4].Value2 = dict[$"{edu}学制"];     //学制

                    #endregion
                }
            }

            #region 新参加工作
            levelCurRow += 1;
            sumWorkSheet2.Cells[levelCurRow, 1].Value2 = dict["序号"];
            sumWorkSheet2.Cells[levelCurRow, 2].Value2 = dict["参加工作时间"];
            sumWorkSheet2.Cells[levelCurRow, 3].Value2 = "新参加工作";
            sumWorkSheet2.Cells[levelCurRow, 4].Value2 = dict["职务类别"];
            sumWorkSheet2.Cells[levelCurRow, 8].Value2 = isNurse;               //特殊岗位
            #endregion

            foreach (string ZJLevel in ZJlevelList)
            {
                if (dict[ZJLevel] != null && dict[ZJLevel] != "")
                {
                    levelCurRow += 1;

                    int levelyear = int.Parse(dict[ZJLevel].Split('.').First());
                    string leveldate = $"{levelyear + 1}.01";

                    sumWorkSheet2.Cells[levelCurRow, 1].Value2 = dict["序号"];
                    sumWorkSheet2.Cells[levelCurRow, 2].Value2 = leveldate;             //变化年月
                    sumWorkSheet2.Cells[levelCurRow, 3].Value2 = "取得资格";            //变化原因
                    sumWorkSheet2.Cells[levelCurRow, 4].Value2 = dict["职务类别"];      //职务类别
                    sumWorkSheet2.Cells[levelCurRow, 5].Value2 = ZJLevel;               //职务
                    sumWorkSheet2.Cells[levelCurRow, 8].Value2 = isNurse;               //特殊岗位
                }
            }

            foreach (string GQLevel in GQlevelList)
            {
                if (dict[GQLevel] != null && dict[GQLevel] != "")
                {
                    levelCurRow += 1;

                    int levelyear = int.Parse(dict[GQLevel].Split('.').First());
                    string leveldate = $"{levelyear + 1}.01";

                    sumWorkSheet2.Cells[levelCurRow, 1].Value2 = dict["序号"];
                    sumWorkSheet2.Cells[levelCurRow, 2].Value2 = leveldate;             //变化年月
                    sumWorkSheet2.Cells[levelCurRow, 3].Value2 = "取得资格";            //变化原因
                    sumWorkSheet2.Cells[levelCurRow, 4].Value2 = dict["职务类别"];      //职务类别
                    sumWorkSheet2.Cells[levelCurRow, 5].Value2 = GQLevel;               //职务
                    sumWorkSheet2.Cells[levelCurRow, 8].Value2 = isNurse;               //特殊岗位
                }
            }

        }


    }
}
