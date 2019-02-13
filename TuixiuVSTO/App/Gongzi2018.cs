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

namespace TuixiuVSTO.App
{
    class Gongzi2018
    {
        public void genSheets()
        {
            int incRow = 9;
            int itemPerPage = 3;
            int curRow = 1;
            string workPath = @"D:\1\";
            string sumFileName = "template.xlsx";

            Excel.Application excelApp = new Excel.Application();
            //Excel.Workbook thisWorkBook = excelApp.ActiveWorkbook;
            //Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(excelApp.ActiveWorkbook.Path + "\\a.xlsx");
            Excel.Workbook thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName);
            Excel.Worksheet thisWorkSheet = thisWorkBook.Worksheets["Sheet1"];

            //workPath = thisWorkBook.Path;
            WriteLine(workPath);


            using (var db = new DBModel())
            {
                var c2018 =
                    db.C2018pay
                    .OrderBy(a => a.科室)
                    .ThenByDescending(a => a.全年收入合计)
                    .ToList();

                int itemCount = c2018.Count;
                //int itemCount = 11;

                foreach (var item in c2018)
                //for (int i = 0; i < itemCount; i++)
                {
                    //var item = c2017[i];
                    WriteLine($"{item.科室}, {item.姓名}, {item.职工号}, {item.身份证号码}, {item.全年收入合计}");
                    curRow = toSheet2(item, thisWorkSheet, curRow, incRow, itemCount, itemPerPage);
                }

            }

            thisWorkBook.SaveAs(Filename: $@"{workPath}done.xlsx");
            thisWorkBook.Close(false);
        }

        public int toSheet2(C2018pay item, Excel.Worksheet thisWorkSheet, int curRow, int incRow, int itemCount, int itemPerPage)
        {
            thisWorkSheet.Cells[2, 2] = item.科室;
            thisWorkSheet.Cells[3, 2] = item.姓名;
            thisWorkSheet.Cells[3, 4] = item.职工号;
            thisWorkSheet.Cells[4, 2] = item.身份证号码;
            thisWorkSheet.Cells[7, 1] = item.全年收入合计;
            thisWorkSheet.Cells[7, 2] = item.工资应发合计;
            thisWorkSheet.Cells[7, 3] = item.全年奖金合计;
            thisWorkSheet.Cells[7, 4] = item.年终奖和补加款;

            string srcStr = "1:9";
            var srcSheet = thisWorkSheet.Range[srcStr];

            curRow = curRow + incRow * itemPerPage;

            if (curRow >= (itemCount + 1) * incRow + 1)
            {
                int pageNum = curRow / (incRow * itemPerPage);
                WriteLine($"分页：{pageNum}");
                curRow = curRow - pageNum * itemPerPage * incRow + incRow;
            }

            string dstStr = $"{curRow}:{curRow + incRow - 1}";
            WriteLine($"src:{srcStr}, dst:{dstStr}");

            var dstSheet = thisWorkSheet.Range[dstStr];

            srcSheet.Copy(dstSheet);

            return curRow;
        }

    }
}
