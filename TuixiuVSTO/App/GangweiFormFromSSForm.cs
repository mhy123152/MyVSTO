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
    class GangweiFormFromSSForm : IDisposable
    {
        public static string workPath = @"D:\1\GangweiFormFromSSForm\";

        //执行年月开始行号
        public static int dataStartNum = 14;

        //执行年月结束行号
        public static int dataEndNum = 33;

        //每页行数, 注意修改！！！
        public static int pageRowCount = 33;

        //退休人员信息的文件名
        public static string dataFileName = "500人历史沿革表.xlsx";
        public static string templateFileName = "事业单位岗位聘用登记表模板.docx";

        public static string resultFileNameString = "岗位聘用登记表";

        public string[] levels;

        Excel.Application excelApp;
        Word.Application wordApp;

        Word.Document templateDocument;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Range ranges;
        Excel.Range searchRanges;

        DateTimeFormatInfo dtFormat;

        string PathHeader;
        string previousName;
        int index = 0;

        public GangweiFormFromSSForm()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();

            dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";

            levels = new[] { "正高级", "副高级", "中级", "助理级", "员级", "见习期、初期", "高级技师", "技师", "高级工", "中级工", "初级工" };

        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            excelApp.Quit();
        }


        public void genForm(int pageNum = 0)
        {
            string sumFilePath = $@"{workPath}{dataFileName}";

            PathHeader = $@"{workPath}result\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            if (!Directory.Exists(PathHeader))
            {
                Directory.CreateDirectory(PathHeader);

            }

            openSheet2(workPath + dataFileName, out thisWorkBook, out thisWorkSheet, "事业在职沿革表");

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


            }
            finally
            {
                thisWorkBook.Close(false);
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

        private void openDoc2(string path, out Word.Document document)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("File cannot found");
                Application.Exit();
            }

            document = wordApp.Documents.Open(path);
        }

        private void doBatch(int pageNum)
        {
            openDoc2(workPath + templateFileName, out templateDocument);


            int firstRow = (pageNum - 1) * pageRowCount + 1;
            int myDataStartNum = firstRow + dataStartNum - 1;
            int myDataEndNum = firstRow + dataEndNum - 1;

            string name = ranges.Cells[firstRow + 2, 2].Text;

            if (!(name == "" || name == null))
            {

                if (previousName != name)
                {
                    previousName = name;
                    index++;
                }

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


                #region 填写表格

                for (int i = myDataStartNum; i <= myDataEndNum; i++)
                {
                    string changeReason = ranges.Cells[i, 1].Text;

                    if (changeReason == "聘用岗位")
                    {
                        string executeDate = ranges.Cells[i, 2].Text;

                        string levelCellValue = ranges.Cells[i, 3].Text;
                        string payCellValue = ranges.Cells[i, 4].Text;

                        foreach (string level in levels)
                        {
                            if (levelCellValue.Contains(level))
                            {
                                levelCellValue = level;
                                break;
                            }
                        }

                        //executeDate.Replace(@".", "-");
                        //executeDate = $"{executeDate}-01";
                        //DateTime executeDateTime = Convert.ToDateTime(executeDate, dtFormat);

                        templateDocument.Variables.Add("姓名", name);
                        templateDocument.Variables.Add("性别", sex);
                        templateDocument.Variables.Add("出生年月", birthDate);
                        templateDocument.Variables.Add("参加工作时间", workDate);
                        templateDocument.Variables.Add("职称", levelCellValue);
                        templateDocument.Variables.Add("聘用时间", executeDate);

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


                        templateDocument.SaveAs2($@"{PathHeader}【{index}】【{name}】【{executeDate}】{resultFileNameString}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);
                        #endregion
                    }


                }


            }


        }

    }
}
