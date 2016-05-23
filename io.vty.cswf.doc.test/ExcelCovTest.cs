using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;
using System.Diagnostics;
using System.Threading;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ExcelCovTest
    {
        [TestInitialize]
        public void init()
        {
            ProcKiller.AddName("EXCEL");
            ProcKiller.StartTimer(300);
        }
        [TestMethod]
        public void TestExcel2img()
        {
            TaskPool.Shared.MaximumConcurrency = 10;
            ExcelCov cov = new ExcelCov("test\\xx.xlsx", "xlsx-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(cov.Result.Count, cov.Result.Files.Count);
            Assert.AreEqual(0, cov.Fails.Count);
        }
        [TestMethod]
        public void TestExcel2img2()
        {
            TaskPool.Shared.MaximumConcurrency = 10;
            for (var i = 0; i < 5; i++)
            {
                ExcelCov cov = new ExcelCov("test\\xx1.xls", "xlsx-d-{0}.jpg");
                cov.Exec();
                cov.PrintFails();
                Assert.AreEqual(cov.Result.Count, cov.Result.Files.Count);
                Assert.AreEqual(0, cov.Fails.Count);
            }
        }
        [TestMethod]
        public void TestExcel2img3()
        {
            TaskPool.Shared.MaximumConcurrency = 10;
            ExcelCov cov = new ExcelCov("test\\xx1.xls", "xlsx1-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(cov.Result.Count, cov.Result.Files.Count);
            Assert.AreEqual(0, cov.Fails.Count);
        }
        [TestMethod]
        public void TestExcel2img4()
        {
            TaskPool.Shared.MaximumConcurrency = 10;
            ExcelCov cov = new ExcelCov("test\\xx2.xls", "xlsx2-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(cov.Result.Count, cov.Result.Files.Count);
            Assert.AreEqual(0, cov.Fails.Count);
        }
        [TestCleanup]
        public void clear()
        {
            ProcKiller.Shared.Running.Clear();
            while (Process.GetProcessesByName("EXCEL").Length > 0)
            {
                Thread.Sleep(1000);
            }
            ExcelCov.Cached.Clear();
            ProcKiller.StopTimer();
        }
    }
}
