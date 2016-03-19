using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class PdfCovTest
    {
        [TestMethod]
        public void TestPdf2img()
        {
            TaskPool.Shared.MaximumConcurrency = 8;
            PdfCov cov = new PdfCov("test\\xx.pdf", "pdf-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(0, cov.Fails.Count);
        }
    }
}
