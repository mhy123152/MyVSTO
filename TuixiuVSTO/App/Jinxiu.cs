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
    class Jinxiu
    {
        string workPath = @"D:\Library\Desktop\2\";
        string sumFileName = "外出进修人员登记表.xlsx";
        string templateFileName = "template.docx";
        string worksheetsName = "国内进修登记";
        int KeyNum = 7;

        Excel.Application excelApp = null;
        Word.Application wordApp = null;
        Excel.Workbook thisWorkBook = null;
        Excel.Worksheet thisWorkSheet = null;
        Word.Document doc = null;

        public Jinxiu()
        {
            excelApp = new Excel.Application();
            wordApp = new Word.Application();


            if (!File.Exists(workPath + templateFileName))
            {
                MessageBox.Show("File cannot found");
                return;
            }

            try
            {
                thisWorkBook = excelApp.Workbooks.Open(workPath + sumFileName, ReadOnly: true);
            }
            catch
            {
                if (thisWorkBook != null)
                {
                    thisWorkBook.Close(false);
                }
                excelApp.Quit();
                MessageBox.Show("Excel Path not found");
                return;
            }

            try
            {
                thisWorkSheet = thisWorkBook.Worksheets[worksheetsName];
            }
            catch
            {
                thisWorkBook.Close(false);
                excelApp.Quit();
                MessageBox.Show("Sheet not found");
                return;
            }


            try
            {
                doc = wordApp.Documents.Open(workPath + templateFileName);
            }
            catch
            {
                if (doc != null)
                {
                    doc.Close();
                }
                MessageBox.Show("Word Path not found");
                return;
            }
        }

        public void jinxiu()
        {
            try
            {
                Excel.Range ranges = thisWorkSheet.UsedRange;

                string PathHeader = $@"{workPath}result\";
                WriteLine(PathHeader);

                if (!Directory.Exists(PathHeader))
                {
                    Directory.CreateDirectory(PathHeader);

                }

                List<model> list = new List<model>();

                for (int i = 3; i <= ranges.Rows.Count; i++)
                {
                    model m = new model();

                    for (int j = 1; j <= KeyNum; j++)
                    {
                        _PropertyInfo p = m.GetType().GetProperty(ranges.Cells[2, j].Text);
                        p.SetValue(m, ranges.Cells[i, j].Text, null);
                    }

                    list.Add(m);
                }


                //根据Date4 经管通知时间 进行分组
                var q = from m in list
                        group m by m.Date4 into g
                        select g;

                //g 为分组之后的集合
                foreach (var g in q)
                {
                    Word.Table table1 = getTable(doc, "table1");
                    Word.Table table2 = getTable(doc, "table2");

                    //g.Key 为去重以后的经管通知时间
                    string date4 = g.Key;
                    WriteLine(date4);

                    //根据每个分组生成word
                    foreach (model m in g)
                    {
                        WriteLine($"{m.Name}: {m.Date3}");
                        if (!(m.Date3 == "" || m.Date3 == null))
                        {
                            //1.进修结束

                            table1.Rows.Add();
                            table1.Rows[table1.Rows.Count - 1].Cells[1].Range.Text = m.Name;
                            table1.Rows[table1.Rows.Count - 1].Cells[2].Range.Text = m.Class;
                            table1.Rows[table1.Rows.Count - 1].Cells[3].Range.Text = m.Place;
                            table1.Rows[table1.Rows.Count - 1].Cells[4].Range.Text = m.Date1;
                            table1.Rows[table1.Rows.Count - 1].Cells[5].Range.Text = m.Date3;
                        }
                        else
                        {
                            //2.外出进修

                            table2.Rows.Add();
                            table2.Rows[table2.Rows.Count - 1].Cells[1].Range.Text = m.Name;
                            table2.Rows[table2.Rows.Count - 1].Cells[2].Range.Text = m.Class;
                            table2.Rows[table2.Rows.Count - 1].Cells[3].Range.Text = m.Place;
                            table2.Rows[table2.Rows.Count - 1].Cells[4].Range.Text = m.Date1;

                        }

                    }

                    doc.Variables.Add("mydate", date4);

                    doc.Fields.Update();

                    doc.SaveAs2($@"{PathHeader}{date4.Substring(0, 8)} 经管补充说明（进修）.docx", FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, LockComments: false, CompatibilityMode: 15);

                    //SaveAs以后，doc就成为了另存为的对象，因此需要关闭
                    doc.Close();

                    //然后重新打开模板文件
                    doc = wordApp.Documents.Open(workPath + templateFileName);

                    //
                    //查看打开的word文档
                    //
                    //foreach(Word.Document d in wordApp.Documents)
                    //{
                    //    WriteLine($"Word.Documents : {d.Name}");
                    //}

                }

            }
            catch (Exception e)
            {
                WriteLine(e.Message);
            }
            finally
            {
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                thisWorkBook.Close(false);
                wordApp.Quit(false);
                excelApp.Quit();
                doc = null;
                thisWorkBook = null;
                GC.Collect();
            }

        }

        public void fillTable(Word.Table table, model m)
        {
            table.Rows.Add();
            table.Rows[table.Rows.Count - 1].Cells[1].Range.Text = m.Name;
            table.Rows[table.Rows.Count - 1].Cells[2].Range.Text = m.Class;
            table.Rows[table.Rows.Count - 1].Cells[3].Range.Text = m.Place;
            table.Rows[table.Rows.Count - 1].Cells[4].Range.Text = m.Date1;
            table.Rows[table.Rows.Count - 1].Cells[5].Range.Text = m.Date2;
            table.Rows[table.Rows.Count - 1].Cells[6].Range.Text = m.Date3;
            table.Rows[table.Rows.Count - 1].Cells[7].Range.Text = m.Date4;
        }

        public static Table getTable(Document doc, String title)
        {
            int totalTables = doc.Tables.Count;
            Microsoft.Office.Interop.Word.Table ret = null;
            for (int i = 1; i <= totalTables; i++)
            {
                if (title.Equals(doc.Tables[i].Title, StringComparison.OrdinalIgnoreCase))
                {
                    ret = doc.Tables[i];
                    break;
                }
            }

            return ret;
        }

        public static Word.Bookmark GetBookmark(Document doc, String bookmarkName)
        {
            // Find bookmark
            Word.Bookmark bookmark = null;
            foreach (Word.Bookmark curBookmark in doc.Bookmarks)
            {
                if (curBookmark.Name.Equals(bookmarkName, StringComparison.OrdinalIgnoreCase))
                {
                    bookmark = curBookmark;
                    break;
                }
            }

            return bookmark;
        }


        public class model
        {
            public string Name { get; set; }
            public string Class { get; set; }
            public string Place { get; set; }
            public string Date1 { get; set; }
            public string Date2 { get; set; }
            public string Date3 { get; set; }
            public string Date4 { get; set; }
        }
    }
}
