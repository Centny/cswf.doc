using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using io.vty.cswf.util;
using System.Collections.Generic;
using System.Collections;

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
            String s = Util.read("test.json");
            Console.WriteLine(s);
            var jres = Json.parse<CovRes>(s);
            Assert.AreEqual(res.Count, jres.Count);
            Assert.AreEqual(res.Files.Count, jres.Files.Count);
        }
        [TestMethod]
        public void TestList()
        {
            List<String> ls = new List<string>(100);
            //ls.Capacity = 100;
            //ls[99] = "xxxx";
            var aa = new ArrayList(100);
            aa[10] = "sdsd";

        }
    }
}
