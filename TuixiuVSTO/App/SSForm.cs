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

namespace TuixiuVSTO.App
{
    class SSForm : IDisposable
    {
        public static string workPath = @"D:\1\ssform\";

        //执行年月开始行号
        public static int dataStartNum = 14;

        //执行年月结束行号
        public static int dataEndNum = 33;

        //每页行数, 注意修改！！！
        public static int pageRowCount = 33;

        //汇总表格的行号
        private int sumFileCurRow = 2;

        //退休人员信息的文件名
        public static string dataFileName = "荆州市中心医院(岗位设置)退休人员沿革表.xlsx";
        public static string sumFileName = "template.xlsx";
        public static string sssumFileName = "历史沿革表汇总数据.xlsx";

        public static string pdfFileName = "荆州市中心医院(岗位设置)退休人员沿革表.pdf";
        public static string resultPdfFileName = "中人历史沿革表(已排序).pdf";

        public string[] names;

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook sumWorkBook;
        Excel.Worksheet sumWorkSheet;

        Excel.Range ranges;
        Excel.Range searchRanges;

        DateTimeFormatInfo dtFormat;

        public SSForm()
        {
            excelApp = new Excel.Application();

            dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";

            names = new[] { "王钟梅", "温万萍", "熊邦琴", "汤小丽", "龚正喜", "文清云", "罗小燕", "刘树玉", "孙芳", "齐琪", "李建英", "龚金萍", "聂品", "文新华", "杨兰珍", "董兰芳", "杨忠珍", "崔云华", "李娟", "王建新", "胡刚", "李家华", "陈明新", "孙其望", "曹安明", "熊文华", "蒋金华", "樊冬兰", "王南", "黄玲", "陈佳君", "段会义", "王辉29", "王争鸣", "熊文凤", "陈明凤", "苏金莲", "刘冬莲", "张辉", "杨有为", "王光亚", "倪新忠", "孙慧慧", "文远梅", "叶红", "刘英", "刘爱蓉", "肖继荣44", "李长庭", "张孝兰", "贾义恒", "卢恒山", "王敏47", "杨志奇", "张先觉", "杨安全", "杨楷", "张家洪", "李厚林", "胡荆江", "龚兰", "何德林", "文汉东", "谢明", "金保山", "黄伟荧", "李欣", "贺国权", "戴鸣", "康宁", "姚长江", "王志贵", "辛棣", "徐正丰", "幸小弘", "皮洁", "王昌富", "杨逸", "谢孝森", "黄谋玲", "胡钢", "李鸿霞", "田方兴", "全为民", "熊怡", "冯向南", "袁丙琴", "宫炫", "时运", "杨德义", "刘道忠", "肖爱华", "赵少华", "万荆发", "冯保亮", "王三元", "陈体文", "张红44", "孟祥存", "叶明华", "甘家铀", "何跃进" };

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

            WriteLine($"{pageNum}:{firstRow}:{name}");

            if (!(name == "" || name == null))
            {
                for (int i = myDataStartNum; i <= myDataEndNum; i++)
                {
                    string executeDate = ranges.Cells[i, 2].Text;
                    string dictValue = ranges.Cells[i, 2].Text;

                    switch (executeDate)
                    {
                        case "2014.10":
                        case "2015.01":
                        case "2016.01":
                        case "2017.01":
                        case "2018.01":
                            executeDate.Replace(@".", "-");
                            executeDate = $"{executeDate}-01";

                            DateTime executeDateTime = Convert.ToDateTime(executeDate, dtFormat);

                            #region 填写表格

                            sumWorkSheet.Cells[sumFileCurRow, 1].Value2 = name;
                            sumWorkSheet.Cells[sumFileCurRow, 2].Value2 = executeDateTime;
                            sumWorkSheet.Cells[sumFileCurRow, 3].Value2 = ranges.Cells[i, 3].Text; //岗位
                            sumWorkSheet.Cells[sumFileCurRow, 4].Value2 = ranges.Cells[i, 4].Text; //薪级
                            sumWorkSheet.Cells[sumFileCurRow, 5].Value2 = ranges.Cells[i, 5].Text; //岗位工资
                            sumWorkSheet.Cells[sumFileCurRow, 6].Value2 = ranges.Cells[i, 6].Text; //薪级工资

                            sumFileCurRow += 1;

                            #endregion

                            break;

                        default:
                            break;
                    }


                    //WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");

                }
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
