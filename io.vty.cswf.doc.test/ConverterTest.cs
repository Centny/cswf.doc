using System.Drawing.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ConverterTest
    {
        int OnProcess(Converter.Res res, int count, string spath)
        {
            res.Files.Add(spath);
            return 1;
        }
        //public string wdir = System.Environment.CurrentDirectory;
        [TestMethod]
        public void TestWord2img()
        {
            Converter.Res res;
            //
            res = Converter.word2img(".\\..\\..\\xx.docx", "docx-{0}.png");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(7, res.Count);
            res.Save("docx0.json");
            //
            res = Converter.word2img(".\\..\\..\\xx.docx", "docx-{0}.png", 0, true, this.OnProcess);
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(7, res.Count);
            res.Save("docx1.json");
        }
        [TestMethod]
        public void TestExcel2pdf()
        {
            Converter.Res res;
            //
            res = Converter.excel2pdf(".\\..\\..\\xx.xlsx", "xlsx-{0}.pdf");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(2, res.Count);
            res.Save("xlsx0.json");
            //
            res = Converter.excel2pdf(".\\..\\..\\xx.xlsx", "xlsx-{0}.pdf", 0, true, this.OnProcess);
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(2, res.Count);
            res.Save("xlsx1.json");
        }
        [TestMethod]
        public void TestPpt2img()
        {
            //new Converter().excel2pdf("C:\\xxx\\xx.xlsx", "C:\\xxx\\xx-{0}.pdf", true);
            Converter.Res res;
            //
            res = Converter.ppt2img(".\\..\\..\\xx.pptx", "ppt-{0}.png");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(1, res.Count);
            res.Save("ppt0.json");
            //
            res = Converter.ppt2img(".\\..\\..\\xx.pptx", "ppt-{0}.png", 0, "png", 0, 0, true, this.OnProcess);
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(1, res.Count);
            res.Save("ppt1.json");
        }

        [TestMethod]
        public void TestExec()
        {
            //new Converter().excel2pdf("C:\\xxx\\xx.xlsx", "C:\\xxx\\xx-{0}.pdf", true);
            //var text=util.Exec("cmd","/c","echo")
            Converter.Res res;
            res = Converter.exec("a", "b");
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(1, res.Count);
            res.Save("exec1.json");

            Converter.Proc proc = new Converter.Proc("..\\..\\echo1.bat", "xx");
            res = Converter.exec("a", "b", 0, false, proc.exec);
            Assert.AreEqual(0, res.Code);
            Assert.AreEqual(1, res.Count);
            res.Save("exec1.json");
        }
    }
}
