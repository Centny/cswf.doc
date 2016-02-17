using System.Drawing.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ConverterTest
    {
        public string wdir = System.Environment.CurrentDirectory;
        [TestMethod]
        public void TestWord2img()
        {
            new Converter().word2img(wdir + "\\..\\..\\xx.docx", wdir + "\\docx-{0}.png",ImageFormat.Png);
        }
        [TestMethod]
        public void TestExcel2pdf()
        {
            new Converter().excel2pdf(wdir + "\\..\\..\\xx.xlsx", wdir + "\\xlsx-{0}.pdf");
        }
        [TestMethod]
        public void TestPpt2img()
        {
            //new Converter().excel2pdf("C:\\xxx\\xx.xlsx", "C:\\xxx\\xx-{0}.pdf", true);
            new Converter().ppt2img(wdir + "\\..\\..\\xx.pptx", wdir + "\\ppt-{0}.png");
        }
    }
}
