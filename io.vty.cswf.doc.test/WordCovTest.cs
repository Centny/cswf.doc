using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;
using System.Diagnostics;
using System.Threading;

namespace io.vty.cswf.doc.test
{
    /// <summary>
    /// Summary description for WordCovTest
    /// </summary>
    [TestClass]
    public class WordCovTest
    {
        public ProcKiller K;
        [TestInitialize]
        public void init()
        {
            ProcKiller.AddName("WINWORD");
            ProcKiller.StartTimer(300);
        }
        [TestMethod]
        public void TestWord2img()
        {
            TaskPool.Shared.MaximumConcurrency = 2;
            WordCov cov = new WordCov("test\\xx.docx", "docx-{0}.jpg");
            cov.Exec();
            cov.PrintFails();
            Assert.AreEqual(0, cov.Fails.Count);
        }
        [TestMethod]
        public void TestWord2img2()
        {
            TaskPool.Shared.MaximumConcurrency = 2;
            for (var i = 0; i < 2; i++)
            {
                WordCov cov = new WordCov("test\\xx.docx", "docx-{0}.jpg");
                cov.Exec();
                cov.PrintFails();
                Assert.AreEqual(0, cov.Fails.Count);
            }
        }
        [TestCleanup]
        public void clear()
        {
            ProcKiller.Shared.Running.Clear();
            while (Process.GetProcessesByName("WINWORD").Length > 0)
            {
                Thread.Sleep(1000);
            }
            WordCov.Cached.Clear();
            ProcKiller.StopTimer();
        }
    }
}
