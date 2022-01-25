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

namespace TuixiuVSTO.App
{
    class Attachment_6
    {

        public static string workPath = @"D:\1\Attachment_6\";

        //起始位置, 注意修改！！！
        public static int startNum = 3;

        //退休人员信息的文件名
        public static string sumFileName = "附件6数据.xlsx";
        public static string dataFileName = "附件6工作经历.xlsx";
        public static string templateFileName = "附件6模板.docx";

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Excel.Workbook dataWorkBook;
        Excel.Worksheet dataWorkSheet;

        Word.Document templateDocument;

        Excel.Range ranges;

        string PathHeader;

        public Attachment_6()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            dataWorkBook.Close(false);
            templateDocument.Close(false);
            excelApp.Quit();
            wordApp.Quit();
        }


        public void genForm(int rowNum = 0)
        {
            PathHeader = $@"{workPath}result\";

            if (Directory.Exists(PathHeader))
            {
                Directory.Delete(PathHeader, true);
            }

            if (!Directory.Exists(PathHeader))
            {
                Directory.CreateDirectory(PathHeader);

            }

            openSheet2(workPath + sumFileName, out thisWorkBook, out thisWorkSheet, "Sheet1");
            openSheet2(workPath + dataFileName, out dataWorkBook, out dataWorkSheet, "Sheet1");

            try
            {
                ranges = thisWorkSheet.UsedRange;

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
                thisWorkBook.Close(false);
                dataWorkBook.Close(false);
                excelApp.Quit();
                wordApp.Quit();
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

        private void doBatch(int rowNum)
        {
            openDoc2(workPath + templateFileName, out templateDocument);

            Dictionary<string, string> dict = new Dictionary<string, string>();

            for (int j = 1; j <= ranges.Columns.Count; j++)
            {
                string dictKey = ranges.Cells[2, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            if (!(dict["id"] == "" || dict["id"] == null))
            {
                /*
                string PathHeader = $@"{workPath}{dict["Name"]}({dict["Class"]})\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }
                */

                #region 填写表格

                foreach (string key in dict.Keys)
                {
                    /*
                    不填写以下数据：
                    nurse，根据情况填写是否护士
                    pay_age, 弃用附件6数据中的缴费月数，根据工作经历来重新计算合计缴费月数
                    */
                    if (key == "nurse" || key == "pay_age")
                    {
                        //Pass
                    }
                    else
                    {
                        templateDocument.Variables.Add(key, dict[key]);
                    }
                }

                //if (dict["tx_class"] == "正常退休")
                //{
                //    templateDocument.Variables.Add("tx_class", "☑");
                //}
                //else
                //{
                //    templateDocument.Variables.Add("tx_class", "□");
                //}

                //根据岗位类型填写
                switch (dict["work_class"])
                {
                    case "事业管理":
                        templateDocument.Variables.Add("work_level_1", dict["work_level"]);
                        templateDocument.Variables.Add("work_paylevel_1", dict["work_paylevel"]);
                        templateDocument.Variables.Add("tx_level_1", dict["tx_level"]);
                        templateDocument.Variables.Add("tx_paylevel_1", dict["tx_paylevel"]);
                        break;
                    case "事业专技":
                        templateDocument.Variables.Add("work_level_2", dict["work_level"]);
                        templateDocument.Variables.Add("work_paylevel_2", dict["work_paylevel"]);
                        templateDocument.Variables.Add("tx_level_2", dict["tx_level"]);
                        templateDocument.Variables.Add("tx_paylevel_2", dict["tx_paylevel"]);
                        break;
                    case "事业工勤":
                        templateDocument.Variables.Add("work_level_3", dict["work_level"]);
                        templateDocument.Variables.Add("work_paylevel_3", dict["work_paylevel"]);
                        templateDocument.Variables.Add("tx_level_3", dict["tx_level"]);
                        templateDocument.Variables.Add("tx_paylevel_3", dict["tx_paylevel"]);
                        break;
                }

                //护士
                if (!(dict["nurse"] == "" || dict["nurse"] == null))
                {
                    templateDocument.Variables.Add("nurse_age", "护龄满20年以上");
                    //templateDocument.Variables.Add("nurse_age_2", $"{int.Parse(dict["nurse_age"]) / 12}年 {int.Parse(dict["nurse_age"]) % 12}个月");
                }

                //职务升降
                if (!(dict["post_change_date"] == "" || dict["post_change_date"] == null))
                {
                    templateDocument.Variables.Add("post_change", "是");

                    if (!(dict["nurse"] == "" || dict["nurse"] == null))
                    {
                        templateDocument.Variables.Add("last_nurse_age", "护龄满20年以上");
                    }

                    //职务升降时，岗位为退休时岗位，薪级为2014年9月薪级
                    switch (dict["work_class"])
                    {
                        case "事业管理":
                            templateDocument.Variables.Add("last_work_level_1", dict["tx_level"]);
                            templateDocument.Variables.Add("last_work_paylevel_1", dict["work_paylevel"]);
                            break;
                        case "事业专技":
                            templateDocument.Variables.Add("last_work_level_2", dict["tx_level"]);
                            templateDocument.Variables.Add("last_work_paylevel_2", dict["work_paylevel"]);
                            break;
                        case "事业工勤":
                            templateDocument.Variables.Add("last_work_level_3", dict["tx_level"]);
                            templateDocument.Variables.Add("last_work_paylevel_3", dict["work_paylevel"]);
                            break;
                    }
                }




                #region 填写工作经历

                Excel.Range dataRange = dataWorkSheet.UsedRange;

                WriteLine($"------------------------------------------------------------------");

                int dataRowNum = 0;
                int pay_age = 0;

                for (int i = 2; i <= dataRange.Rows.Count; i++)
                {
                    if (dataRange.Cells[i, 1].Text == dict["name"])
                    {
                        Write($"DataRange:");
                        Write($"{dataRange.Cells[i, 2].Text}, ");
                        Write($"{dataRange.Cells[i, 3].Text}, ");
                        Write($"{dataRange.Cells[i, 4].Text}, ");
                        Write($"{dataRange.Cells[i, 5].Text}, ");
                        Write($"{dataRange.Cells[i, 6].Text}, ");
                        Write($"{dataRange.Cells[i, 7].Text}, ");
                        WriteLine("");

                        dataRowNum++;

                        for (int j = 1; j <= 5; j++)
                        {
                            templateDocument.Variables.Add($"r{dataRowNum}c{j}", dataRange.Cells[i, 1 + j].Text);
                        }

                        //根据工作经历来计算合计缴费月数
                        pay_age += int.Parse(dataRange.Cells[i, 6].Text);

                    }
                }

                WriteLine($"------------------------------------------------------------------");

                templateDocument.Variables.Add("pay_age", pay_age);

                #endregion


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


                #endregion

                templateDocument.SaveAs2($@"{PathHeader}【{dict["tx_date"]}】【{dict["name"]}】附件6.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                templateDocument.Close(SaveChanges: false);

            }
        }
    }
}
