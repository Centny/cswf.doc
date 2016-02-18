using System.Drawing.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ConverterTest
    {
        //public string wdir = System.Environment.CurrentDirectory;
        [TestMethod]
        public void TestWord2img()
        {
            var res = Converter.word2img(".\\..\\..\\xx.docx", "docx-{0}.png");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(7, res.Count);
            res.Save("docx.json");
        }
        [TestMethod]
        public void TestExcel2pdf()
        {
            var res = Converter.excel2pdf(".\\..\\..\\xx.xlsx", "xlsx-{0}.pdf");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(2, res.Count);
            res.Save("xlsx.json");
        }
        [TestMethod]
        public void TestPpt2img()
        {
            //new Converter().excel2pdf("C:\\xxx\\xx.xlsx", "C:\\xxx\\xx-{0}.pdf", true);
            var res = Converter.ppt2img(".\\..\\..\\xx.pptx", "ppt-{0}.png");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(1, res.Count);
            res.Save("ppt.json");
        }
    }
}
