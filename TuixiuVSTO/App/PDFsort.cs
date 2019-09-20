using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using System.Drawing;
using Syncfusion.Pdf.Grid;
using System.Data;
using Syncfusion.Pdf.Parsing;

namespace TuixiuVSTO.App
{
    class PDFsort
    {
        public static string workPath = @"D:\1\";
        public static string pdfFileName = "combine.pdf";

        public int[] pages;


        public PDFsort()
        {
            pages = new[] { 175, 42, 872, 866, 921, 74, 782, 110, 192, 835, 923, 240, 829, 887, 18, 184, 833, 104, 100, 230, 807, 450, 109, 881, 76, 229, 60, 123, 131, 56, 5, 106, 886, 908, 176, 156, 690, 846, 813, 169, 138, 155, 134, 53, 19, 212, 221, 125, 911, 47, 851, 31, 164, 1, 811, 96, 215, 148, 920, 822, 151, 906, 858, 17, 930, 898, 140, 218, 95, 793, 854, 917, 158, 137, 840, 899, 162, 181, 912, 57, 220, 824, 933, 837, 108, 61, 67, 932, 838, 790, 25, 204, 871, 855, 32, 239, 852, 910, 869, 55, 142, 50, 823, 784, 805, 48, 213, 895, 179, 90, 233, 130, 24, 54, 868, 860, 227, 202, 841, 290, 927, 580, 71, 794, 152, 178, 41, 37, 660, 69, 170, 186, 103, 816, 850, 812, 780, 280, 65, 250, 14, 896, 804, 859, 9, 63, 75, 879, 193, 901, 797, 926, 876, 26, 857, 99, 45, 191, 808, 166, 924, 787, 116, 182, 925, 919, 799, 320, 121, 905, 93, 113, 44, 885, 10, 154, 270, 198, 163, 848, 171, 785, 902, 135, 52, 863, 122, 180, 107, 187, 77, 167, 16, 826, 928, 51, 7, 226, 115, 118, 831, 830, 897, 328, 200, 235, 133, 796, 94, 79, 922, 91, 177, 864, 145, 2, 132, 80, 58, 62, 828, 6, 231, 832, 20, 73, 13, 878, 865, 904, 803, 420, 38, 3, 810, 209, 882, 888, 117, 40, 909, 12, 68, 188, 150, 845, 22, 114, 779, 821, 34, 197, 194, 791, 4, 834, 27, 112, 85, 781, 189, 146, 216, 330, 844, 153, 64, 174, 249, 847, 815, 66, 893, 46, 160, 861, 168, 798, 849, 795, 98, 818, 929, 111, 825, 119, 173, 196, 916, 222, 11, 101, 23, 329, 82, 783, 219, 225, 105, 875, 127, 15, 39, 214, 238, 210, 183, 206, 800, 201, 843, 36, 139, 124, 143, 43, 853, 129, 862, 817, 778, 207, 89, 224, 232, 890, 889, 319, 29, 894, 120, 918, 903, 877, 33, 900, 236, 836, 913, 35, 83, 870, 172, 21, 935, 827, 208, 190, 931, 786, 72, 873, 289, 161, 199, 205, 78, 814, 856, 237, 907, 97, 149, 891, 165, 789, 792, 802, 211, 801, 880, 806, 788, 867, 203, 892, 260, 30, 874, 217, 234, 147, 884, 223, 59, 84, 883, 8, 842, 915, 820, 144, 819, 159, 86, 269, 92, 228, 839, 87, 195, 28, 102, 126, 128, 934, 141, 809, 185, 70, 914, 81, 88, 49, 157, 136 };

        }

        public void CreatePDFwithTable()
        {
            //Create a new PDF document.
            PdfDocument doc = new PdfDocument();
            //Add a page.
            PdfPage page = doc.Pages.Add();
            //Create a PdfGrid.
            PdfGrid pdfGrid = new PdfGrid();
            //Create a DataTable.
            DataTable dataTable = new DataTable();
            //Add columns to the DataTable
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Name");
            //Add rows to the DataTable.
            dataTable.Rows.Add(new object[] { "E01", "Clay" });
            dataTable.Rows.Add(new object[] { "E02", "Thomas" });
            dataTable.Rows.Add(new object[] { "E03", "Andrew" });
            dataTable.Rows.Add(new object[] { "E04", "Paul" });
            dataTable.Rows.Add(new object[] { "E05", "Gary" });
            //Assign data source.
            pdfGrid.DataSource = dataTable;
            //Draw grid to the page of PDF document.
            pdfGrid.Draw(page, new PointF(10, 10));
            //Save the document.
            doc.Save("Output.pdf");
            //close the document
            doc.Close(true);
        }

        public void MergePDF()
        {
            //Loads document

            PdfLoadedDocument combinePDF = new PdfLoadedDocument(workPath + pdfFileName);

            PdfDocument newPDF = new PdfDocument();

            //Imports the page at 1 from the lDoc
            foreach (int pageNum in pages)
            {
                newPDF.ImportPage(combinePDF, pageNum-1);
            }



            //Saves the document

            newPDF.Save("new.pdf");

            //Closes the document

            newPDF.Close(true);

            combinePDF.Close(true);
        }
    }
}
