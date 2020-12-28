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
    class GangweiForm : IDisposable
    {
        #region 需要修改的变量
        public static string resultFileNameString = "岗位聘用登记表";
        public static string sumFileName = "500人信息表-仅保留数值.xlsx";
        public static string templateFileName = "事业单位岗位聘用登记表模板.docx";
        #endregion

        #region 固定变量
        public static string workPath = @"D:\1\gangweiform\";

        //起始位置, 注意修改！！！
        public static int startNum = 2;

        public int keyNum;

        #endregion

        Excel.Application excelApp;
        Word.Application wordApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet thisWorkSheet;

        Word.Document templateDocument;

        Excel.Range ranges;

        string PathHeader;

        public GangweiForm()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            templateDocument.Close(false);
            excelApp.Quit();
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

            }
            finally
            {
                thisWorkBook.Close(false);
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


            Dictionary<string, string> dict = new Dictionary<string, string>();
            List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();

            for (int j = 1; j <= keyNum; j++)
            {
                string dictKey = ranges.Cells[startNum - 1, j].Text;
                string dictValue = ranges.Cells[rowNum, j].Text;
                WriteLine($"dictKey:{dictKey}, dictValue:{dictValue}");
                dict.Add(dictKey, dictValue);
            }

            //if (dict["副高级"] != null && dict["副高级"] != "") { dict.Add("职称", "副高级"); dict.Add("聘用时间", dict["副高级"]); dict.Add("专技等级", "七"); }
            //else if (dict["中级"] != null && dict["中级"] != "") { dict.Add("职称", "中级"); dict.Add("聘用时间", dict["中级"]); dict.Add("专技等级", "十"); }
            //else if (dict["助理级"] != null && dict["助理级"] != "") { dict.Add("职称", "助理级"); dict.Add("聘用时间", dict["助理级"]); dict.Add("专技等级", "十二"); }
            //else if (dict["员级"] != null && dict["员级"] != "") { dict.Add("职称", "员级"); dict.Add("聘用时间", dict["员级"]); dict.Add("专技等级", "十三"); }
            //else if (dict["高级技师"] != null && dict["高级技师"] != "") { dict.Add("职称", "高级技师"); dict.Add("聘用时间", dict["高级技师"]); dict.Add("工勤等级", "一"); }
            //else if (dict["技师"] != null && dict["技师"] != "") { dict.Add("职称", "技师"); dict.Add("聘用时间", dict["技师"]); dict.Add("工勤等级", "二"); }
            //else if (dict["高级工"] != null && dict["高级工"] != "") { dict.Add("职称", "高级工"); dict.Add("聘用时间", dict["高级工"]); dict.Add("工勤等级", "三"); }
            //else if (dict["中级工"] != null && dict["中级工"] != "") { dict.Add("职称", "中级工"); dict.Add("聘用时间", dict["中级工"]); dict.Add("工勤等级", "四"); }
            //else if (dict["初级工"] != null && dict["初级工"] != "") { dict.Add("职称", "初级工"); dict.Add("聘用时间", dict["初级工"]); dict.Add("工勤等级", "五"); }



            if (!(dict["id"] == "" || dict["id"] == null))
            {
                Dictionary<string, string> tempDict;
                if (dict["副高级"] != null && dict["副高级"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "副高级"); tempDict.Add("聘用时间", dict["副高级"]); tempDict.Add("专技等级", "七"); }
                if (dict["中级"] != null && dict["中级"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "中级"); tempDict.Add("聘用时间", dict["中级"]); tempDict.Add("专技等级", "十"); }
                if (dict["助理级"] != null && dict["助理级"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "助理级"); tempDict.Add("聘用时间", dict["助理级"]); tempDict.Add("专技等级", "十二"); }
                if (dict["员级"] != null && dict["员级"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "员级"); tempDict.Add("聘用时间", dict["员级"]); tempDict.Add("专技等级", "十三"); }
                if (dict["高级技师"] != null && dict["高级技师"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "高级技师"); tempDict.Add("聘用时间", dict["高级技师"]); tempDict.Add("工勤等级", "一"); }
                if (dict["技师"] != null && dict["技师"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "技师"); tempDict.Add("聘用时间", dict["技师"]); tempDict.Add("工勤等级", "二"); }
                if (dict["高级工"] != null && dict["高级工"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "高级工"); tempDict.Add("聘用时间", dict["高级工"]); tempDict.Add("工勤等级", "三"); }
                if (dict["中级工"] != null && dict["中级工"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "中级工"); tempDict.Add("聘用时间", dict["中级工"]); tempDict.Add("工勤等级", "四"); }
                if (dict["初级工"] != null && dict["初级工"] != "") { tempDict = new Dictionary<string, string>(); list.Add(tempDict); tempDict.Add("职称", "初级工"); tempDict.Add("聘用时间", dict["初级工"]); tempDict.Add("工勤等级", "五"); }

                foreach (Dictionary<string, string> feDict in list)
                {
                    openDoc2(workPath + templateFileName, out templateDocument);

                    WriteLine($"dictKey:{"职称"}, dictValue:{feDict["职称"]}");
                    WriteLine($"dictKey:{"聘用时间"}, dictValue:{feDict["聘用时间"]}");

                    //排除聘用时间为间断的单元格
                    if (feDict["聘用时间"].Contains('-'))
                    {
                        //转换时间格式为：yyyy年mm月
                        feDict.Add("聘用时间2", $"{feDict["聘用时间"].Split('-')[0]}年{feDict["聘用时间"].Split('-')[1]}月");
                        //feDict.Add("落款时间", "2020年11月");
                        feDict.Add("落款时间", feDict["聘用时间2"]);

                        #region 填写表格

                        foreach (string key in dict.Keys)
                        {
                            templateDocument.Variables.Add(key, dict[key]);
                        }

                        foreach (string key in feDict.Keys)
                        {
                            templateDocument.Variables.Add(key, feDict[key]);
                        }


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

                        //templateDocument.SaveAs2($@"{PathHeader}【{string.Format("{0:d4}", dict["id"])}】{resultFileNameString}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                        templateDocument.SaveAs2($@"{PathHeader}【{dict["标识"]}】【{dict["姓名"]}】【{feDict["聘用时间"]}】{resultFileNameString}.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                        templateDocument.Close(SaveChanges: false);
                    }

                    
                }



            }
        }
    }
}
