using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class CovResTest
    {
        [TestMethod]
        public void TestCovRes()
        {
            var res = new CovRes("abc.doc");
            res.Count += 1;
            res.Files.Add("ss/abc-0.jpg");
            res.Trim("ss/");
            res.Save("test.json");
            String s=Util.read("test.json");
            var jres = Json.parse<CovRes>(s);
            Assert.AreEqual(res.Count, jres.Count);
            Assert.AreEqual(res.Files.Count, jres.Files.Count);
        }
    }
}
