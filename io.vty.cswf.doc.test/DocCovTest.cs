using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;
using System.Threading;
using System.Collections.Generic;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class DocCovTest
    {
        public class DocCovT : DocCov
        {
            public object done;
            public float rate;
            public bool error;
            public bool do_err;
            public DocCovT(string name, FCfg cfg) : base(name, cfg)
            {

            }

            protected override void SendDone(object args)
            {
                if (this.error)
                {
                    this.do_err = true;
                    throw new Exception("error");
                }
                this.done = args;
            }
            public override void NotifyProc(string tid, float rate)
            {
                if (this.error)
                {
                    this.do_err = true;
                    throw new Exception("error");
                }
                if (this.rate > rate)
                {
                    throw new Exception("rate fail");
                }
                this.rate = rate;
            }

            public void runFailSupport()
            {
                base.RunSupported("sss", SupportedL.None, null, null, "", "");
            }
        }
        [TestMethod]
        public void TestDoWord()
        {
            TaskPool.Shared.MaximumConcurrency = 3;
            FCfg cfg = new FCfg();
            DocCovT cov;
            cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a1", cfg, "Word test\\xx.docx docx_00-{0}.jpg 768 1024");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            var data_ = cov.done as IDictionary<string, object>;
            var data = new Dict(data_);
            Assert.AreEqual(0, data.Val("code", -1));
            Assert.AreEqual("a1", data.Val("tid", ""));
            var res = data["data"] as CovRes;
            Assert.AreEqual(true, res.Count > 0 && res.Count == res.Files.Count);
            Assert.AreNotEqual(0, cov.rate);
            //
            cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a2", cfg, "Word test\\xx.docxx docx_00-{0}.jpg 768 1024");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            data_ = cov.done as IDictionary<string, object>;
            data = new Dict(data_);
            Assert.AreNotEqual(0, data.Val("code", 0));
            Assert.AreEqual("a2", data.Val("tid", ""));
            var err = data["err"] as String;
            Assert.AreNotEqual(0, err.Length);
            //
            //
            cov = new DocCovT("DocCov", cfg);
            cov.error = true;
            cov.DoCmd("a2", cfg, "Word test\\xx.docx docx_01-{0}.jpg 768 1024");
            while (!cov.do_err)
            {
                Thread.Sleep(500);
            }
        }
        [TestMethod]
        public void TestDoPowerPoint()
        {
            FCfg cfg = new FCfg();
            DocCovT cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a1", cfg, "PowerPoint test\\xx.pptx pptx_00-{0}.jpg");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            var data_ = cov.done as IDictionary<string, object>;
            var data = new Dict(data_);
            Assert.AreEqual(0, data.Val("code", -1));
            Assert.AreEqual("a1", data.Val("tid", ""));
            var res = data["data"] as CovRes;
            Assert.AreEqual(true, res.Count > 0 && res.Count == res.Files.Count);
            Assert.AreNotEqual(0, cov.rate);
            //
            cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a2", cfg, "PowerPoint test\\xx.pptxx pptxx_00-{0}.jpg 768 1024");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            data_ = cov.done as IDictionary<string, object>;
            data = new Dict(data_);
            Assert.AreNotEqual(0, data.Val("code", 0));
            Assert.AreEqual("a2", data.Val("tid", ""));
            var err = data["err"] as String;
            Assert.AreNotEqual(0, err.Length);
        }

        [TestMethod]
        public void TestDoExcel()
        {
            FCfg cfg = new FCfg();
            DocCovT cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a1", cfg, "Excel test\\xx.xlsx xlsx_00-{0}.jpg");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            var data_ = cov.done as IDictionary<string, object>;
            var data = new Dict(data_);
            Assert.AreEqual(0, data.Val("code", -1));
            Assert.AreEqual("a1", data.Val("tid", ""));
            var res = data["data"] as CovRes;
            Assert.AreEqual(true, res.Count > 0 && res.Count == res.Files.Count);
            Assert.AreNotEqual(0, cov.rate);
            //
            cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a2", cfg, "Excel test\\xx.xlsxx xlsxx_00-{0}.jpg 768 1024");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            data_ = cov.done as IDictionary<string, object>;
            data = new Dict(data_);
            Assert.AreNotEqual(0, data.Val("code", 0));
            Assert.AreEqual("a2", data.Val("tid", ""));
            var err = data["err"] as String;
            Assert.AreNotEqual(0, err.Length);
        }
        [TestMethod]
        public void TestNormal()
        {
            FCfg cfg = new FCfg();
            DocCovT cov = new DocCovT("DocCov", cfg);
            cov.DoCmd("a1", cfg, "test\\dtm_json.bat");
            while (cov.done == null)
            {
                Thread.Sleep(500);
            }
            var data_ = cov.done as IDictionary<string, object>;
            var data = new Dict(data_);
            Assert.AreEqual(0, data.Val("code", -1));
            Assert.AreEqual("a1", data.Val("tid", ""));
            var res = data["data"] as IDictionary<string, object>;
            Assert.AreEqual(true, res.Count > 0);
        }

        [TestMethod]
        public void TestErr()
        {
            DocCovT cov;
            FCfg cfg = new FCfg();
            cov = new DocCovT("DocCov", cfg);
            try
            {
                cov.DoCmd("a2", cfg, "");
                Assert.Fail();
            }
            catch (Exception)
            {

            }
            cov = new DocCovT("DocCov", cfg);
            try
            {
                cov.builder();
                Assert.Fail();
            }
            catch (Exception)
            {

            }
            new DocCov("DocCov", cfg, () =>
            {
                return null;
            });
            cov = new DocCovT("DocCov", cfg);
            try
            {
                cov.DoCmd("a2", cfg, "Word");
                Assert.Fail();
            }
            catch (Exception)
            {

            }
            cov = new DocCovT("DocCov", cfg);
            try
            {
                cov.runFailSupport();
                Assert.Fail();
            }
            catch (Exception)
            {

            }
        }
        [TestCleanup]
        public void clear()
        {
            WordCov.Cached.Clear();
            PowerPointCov.Cached.Clear();
            ExcelCov.Cached.Clear();
        }
    }
}
