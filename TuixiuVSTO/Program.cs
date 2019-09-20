using TuixiuVSTO.App;

namespace TuixiuVSTO
{
    class Program
    {
        [System.STAThread]
        static void Main(string[] args)
        {
            //Tuixiu tx = new Tuixiu();
            //tx.tuixiu();

            //Zhenggao zg = new Zhenggao();
            //zg.zhenggao2018();

            //Gongzi gz = new Gongzi();
            //gz.genSheets();

            //Jinxiu jx = new Jinxiu();
            //jx.jinxiu();

            //Gongzi2018 gongzi2018 = new Gongzi2018();
            //gongzi2018.genSheets();

            //TuixiuForm txForm = new TuixiuForm();
            //txForm.genForm(new int[] { 99, 102 });
            //txForm.genForm(dateBefore: DateTime.MaxValue, dateAfter: new DateTime(2019,7,1));

            //XuegongLunzhuan xglz = new XuegongLunzhuan();
            //xglz.genForm();

            //Attachment_6 attachment_6 = new Attachment_6();
            //attachment_6.genForm();

            //PDFsort pdfsort = new PDFsort();
            //pdfsort.MergePDF();

            SSForm ssForm = new SSForm();
            //ssForm.genForm();
            ssForm.myExcelToPDFConversion();

        }

    }
}
