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

namespace TuixiuVSTO.App
{
    class XuegongLunzhuan
    {
        public static string workPath = @"D:\Library\Desktop\1\XuegongLunzhuan\";

        //字段数量
        public static int keyNum = 1;

        //起始位置, 注意修改！！！
        public static int startNum = 1;

        //退休人员信息的文件名
        public static string dataFileName = "test.xlsx";
        public static string resultFileName = "result.xlsx";

        Excel.Application excelApp;

        Excel.Workbook thisWorkBook;
        Excel.Worksheet nameWorkSheet;
        Excel.Worksheet classWorkSheet;

        public XuegongLunzhuan()
        {
            excelApp = new Excel.Application();
        }

        public void Dispose()
        {
            thisWorkBook.Close(false);
            excelApp.Quit();
        }


        public void genForm(int rowNum = 0)
        {

            string path = workPath + dataFileName;

            if (!File.Exists(path))
            {
                MessageBox.Show("File cannot found");
                Application.Exit();
            }

            thisWorkBook = excelApp.Workbooks.Open(path);

            nameWorkSheet = thisWorkBook.Worksheets["Name"];
            classWorkSheet = thisWorkBook.Worksheets["Class"];

            Excel.Range nameRange = nameWorkSheet.UsedRange;
            Excel.Range classRange = classWorkSheet.UsedRange;

            List<string> nameList = new List<string>();
            List<string> classList = new List<string>();

            try
            {
                for (int i = startNum; i <= nameRange.Rows.Count; i++)
                {
                    nameList.Add(nameRange.Cells[i, 1].Text);
                }

                for (int i = startNum; i <= classRange.Rows.Count; i++)
                {
                    classList.Add(classRange.Cells[i, 1].Text);
                }

                for(int i=0; i<nameList.Count; i++)
                {
                    WriteLine(nameList[i]);
                }

                for (int i = 0; i < classList.Count; i++)
                {
                    WriteLine(classList[i]);
                }



            }
            finally
            {
                thisWorkBook.Close(false);
                excelApp.Quit();
            }

        }

    }
}
