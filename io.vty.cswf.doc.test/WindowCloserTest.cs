using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class WindowCloserTest
    {
        [TestMethod]
        public void TestSendClose()
        {
            new WindowCloser().SendClose();
        }
    }
}
