using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ImgCovTest
    {
        [TestMethod]
        public void TestImgCov()
        {
            ImgCov cov = new ImgCov("test\\xx.png", "img1.jpg");
            cov.Exec();
            cov.Dispose();
            Assert.AreEqual(0, cov.Fails.Count);
            cov = new ImgCov("test\\xx.pxng", "img1.jpg");
            cov.Exec();
            cov.Dispose();
            Assert.AreNotEqual(0, cov.Fails.Count);
        }
        [TestMethod]
        public void TestImgCov2()
        {
            ImgCov cov = new ImgCov("test\\xx.jpg", "img2.jpg");
            cov.Exec();
            cov.Dispose();
            Assert.AreEqual(0, cov.Fails.Count);
        }

        [TestMethod]
        public void TestImgCov3()
        {
            ImgCov cov = new ImgCov("F:/sdata_i/www/u_57330b08bc9a34219700001c.jpg","F:/sdata_o/www/u_57330b08bc9a34219700001c.jpg");
            cov.Exec();
            cov.Dispose();
            //Console.WriteLine(cov.Fails[0]);
            Assert.AreEqual(0, cov.Fails.Count);
        }
    }
}
