using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class PptxCovTest
    {
        [TestMethod]
        public void TestPptx2img()
        {
            TaskPool.Shared.MaximumConcurrency = 3;
            PowerPointCov cov = new PowerPointCov("test\\xx.pptx", "pptx-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(0, cov.Fails.Count);
        }
        [TestMethod]
        public void TestPptx2img2()
        {
            TaskPool.Shared.MaximumConcurrency = 3;
            for (var i = 0; i < 10; i++)
            {
                PowerPointCov cov = new PowerPointCov("test\\xx.pptx", "pptx-{0}.jpg");
                cov.Exec();
                cov.PrintFails();
                Assert.AreEqual(0, cov.Fails.Count);
            }
        }
        [TestCleanup]
        public void clear()
        {
            PowerPointCov.Cached.Clear();
        }
    }
}
