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
    class SSForm500 : IDisposable
    {
        //public static string workPath = @"D:\1\ssform500\";
        public static string workPath = @"C:\Users\Michael\OneDrive\0Working\500人养老保险\ssform500\";

        //执行年月开始行号
        public static int dataStartNum = 14;

        //执行年月结束行号
        public static int dataEndNum = 33;

        //每页行数, 注意修改！！！
        public static int pageRowCount = 33;

        //汇总表格的标题行号
        private int sumFileCurRow = 1;

        //退休人员信息的文件名
        public static string dataFileName = "500人历史沿革表.xlsx";
        public static string sumFileName = "template.xlsx";
        public static string sssumFileName = "500人历史沿革表汇总数据.xlsx";

        public static string pdfFileName = "荆州市中心医院(岗位设置)退休人员沿革表.pdf";
        public static string resultPdfFileName = "中人历史沿革表(已排序).pdf";

        public string[] names;
        public string[] levels;

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

        public SSForm500()
        {
            excelApp = new Excel.Application();

            dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";

            levels = new[] { "正高级", "副高级", "中级", "助理级", "员级", "见习期、初期", "高级技师", "技师", "高级工", "中级工", "初级工" };

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
            openSheet2(workPath + sumFileName, out sumWorkBook2, out sumWorkSheet2, "Sheet2");

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
                    sumFileCurRow += 1;  //当前行号加一

                    string sex = ranges.Cells[firstRow + 2, 4].Text;
                    string birthDateCellValue = ranges.Cells[firstRow + 4, 2].Text;
                    string eduCellValue = ranges.Cells[firstRow + 6, 2].Text;
                    string workDateCellValue = ranges.Cells[firstRow + 2, 8].Text;

                    DateTime birthDate = Convert.ToDateTime(birthDateCellValue, dtFormat);

                    string edu = "";
                    if (!(eduCellValue == "" || eduCellValue == null))
                    {
                        //string edu = Regex.Match(eduCellValue, @"\.[0-9][0-9][\S]+?;$").Value;  //正则表达式判断，lazy模式不成功
                        string eduAndDate = eduCellValue.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Last();
                        edu = eduAndDate.Substring(7);
                    }

                    DateTime workDate = new DateTime();
                    if (workDateCellValue.Contains("新参加工作"))
                    {
                        string workDateText = workDateCellValue.Substring(0, 7);
                        workDate = Convert.ToDateTime(workDateText, dtFormat);
                    }

                    WriteLine($"{pageNum}:{firstRow}:{name}:{sex}:{birthDate:yyyy-MM}:{edu}:{workDate:yyyy-MM}");


                    #region 填写表格（基本信息）

                    sumWorkSheet.Cells[sumFileCurRow, 1].Value2 = pageNum;      //页码号
                    sumWorkSheet.Cells[sumFileCurRow, 2].Value2 = name;         //姓名
                    sumWorkSheet.Cells[sumFileCurRow, 3].Value2 = sex;          //性别
                    sumWorkSheet.Cells[sumFileCurRow, 4].Value2 = birthDate;    //出生年月
                    sumWorkSheet.Cells[sumFileCurRow, 5].Value2 = workDate;     //参加工作时间
                    sumWorkSheet.Cells[sumFileCurRow, 6].Value2 = edu;          //最高学历

                    #endregion

                    sumWorkSheet2.Cells[sumFileCurRow, 1].Value2 = pageNum;      //页码号
                    sumWorkSheet2.Cells[sumFileCurRow, 2].Value2 = name;         //姓名
                }

                #region 填写表格（工资信息）

                for (int i = myDataStartNum; i <= myDataEndNum; i++)
                {
                    string executeDate = ranges.Cells[i, 2].Text;

                    string levelCellValue = ranges.Cells[i, 3].Text;
                    string payCellValue = ranges.Cells[i, 4].Text;

                    if (levelCellValue.Contains("中级工"))
                    {
                        levelCellValue = "中级工";
                    }
                    else
                    {
                        foreach (string level in levels)
                        {

                            if (levelCellValue.Contains(level))
                            {
                                levelCellValue = level;
                                break;
                            }
                        }
                    }

                    //executeDate.Replace(@".", "-");
                    //executeDate = $"{executeDate}-01";
                    //DateTime executeDateTime = Convert.ToDateTime(executeDate, dtFormat);

                    switch (executeDate)
                    {
                        case "2014.09":
                        case "2014.10":
                            sumWorkSheet2.Cells[sumFileCurRow, 3].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 4].Value2 = payCellValue; //薪级
                            break;
                        case "2015.01":
                        case "2015.02":
                        case "2015.03":
                        case "2015.04":
                        case "2015.05":
                        case "2015.06":
                        case "2015.07":
                        case "2015.08":
                        case "2015.09":
                        case "2015.10":
                        case "2015.11":
                        case "2015.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 5].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 6].Value2 = payCellValue; //薪级
                            break;
                        case "2016.01":
                        case "2016.02":
                        case "2016.03":
                        case "2016.04":
                        case "2016.05":
                        case "2016.06":
                        case "2016.07":
                        case "2016.08":
                        case "2016.09":
                        case "2016.10":
                        case "2016.11":
                        case "2016.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 7].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 8].Value2 = payCellValue; //薪级
                            break;
                        case "2017.01":
                        case "2017.02":
                        case "2017.03":
                        case "2017.04":
                        case "2017.05":
                        case "2017.06":
                        case "2017.07":
                        case "2017.08":
                        case "2017.09":
                        case "2017.10":
                        case "2017.11":
                        case "2017.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 9].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 10].Value2 = payCellValue; //薪级
                            break;
                        case "2018.01":
                        case "2018.02":
                        case "2018.03":
                        case "2018.04":
                        case "2018.05":
                        case "2018.06":
                        case "2018.07":
                        case "2018.08":
                        case "2018.09":
                        case "2018.10":
                        case "2018.11":
                        case "2018.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 11].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 12].Value2 = payCellValue; //薪级
                            break;
                        case "2019.01":
                        case "2019.02":
                        case "2019.03":
                        case "2019.04":
                        case "2019.05":
                        case "2019.06":
                        case "2019.07":
                        case "2019.08":
                        case "2019.09":
                        case "2019.10":
                        case "2019.11":
                        case "2019.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 13].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 14].Value2 = payCellValue; //薪级
                            break;
                        case "2020.01":
                        case "2020.02":
                        case "2020.03":
                        case "2020.04":
                        case "2020.05":
                        case "2020.06":
                        case "2020.07":
                        case "2020.08":
                        case "2020.09":
                        case "2020.10":
                        case "2020.11":
                        case "2020.12":
                            sumWorkSheet2.Cells[sumFileCurRow, 15].Value2 = levelCellValue; //岗位
                            sumWorkSheet2.Cells[sumFileCurRow, 16].Value2 = payCellValue; //薪级
                            break;
                        default:
                            break;
                    }
                    #endregion

                }

                previousName = name;

            }

        }

        public void ExcelToPDFConversion()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Open($@"{workPath}{dataFileName}", ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets["事业在职沿革表"];

                //convert the sheet to PDF
                ExcelToPdfConverter converter = new ExcelToPdfConverter(sheet);

                PdfDocument pdfDocument = new PdfDocument();
                pdfDocument = converter.Convert();
                pdfDocument.Save($@"{workPath}ExcelToPDF.pdf");
            }
        }

        public void myExcelToPDFConversion()
        {
            //Loads document
            PdfLoadedDocument dataPDF = new PdfLoadedDocument(workPath + pdfFileName);

            PdfDocument newPDF = new PdfDocument();

            //创建一个空白页面,并放在第一页,避免水印
            newPDF.Pages.Add();

            openSheet2(workPath + dataFileName, out thisWorkBook, out thisWorkSheet, "事业在职沿革表");

            try
            {
                ranges = thisWorkSheet.UsedRange;
                searchRanges = thisWorkSheet.UsedRange.Range["B:B"];

                //for (int page = 1; page <= ranges.Rows.Count / pageRowCount; page++)
                //{
                //    doBatch(page);
                //}

                //Imports the page at 1 from the lDoc
                foreach (string name in names)
                {
                    Excel.Range currentFind = null;

                    currentFind = searchRanges.Find(
                        name,
                        Type.Missing,
                        Excel.XlFindLookIn.xlValues,
                        Excel.XlLookAt.xlPart,
                        Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlNext,
                        false,
                        Type.Missing,
                        Type.Missing);

                    if (currentFind != null)
                    {
                        int pageNum = (currentFind.Row - 3) / pageRowCount + 1;

                        WriteLine($"{name}:{pageNum}");

                        newPDF.ImportPage(dataPDF, pageNum - 1);

                    }

                }

            }
            finally
            {
                thisWorkBook.Close(false);
                excelApp.Quit();
            }

            //Saves the document
            newPDF.Save($@"{workPath}{resultPdfFileName}");

            //Closes the document
            newPDF.Close(true);

            dataPDF.Close(true);
        }


    }
}
