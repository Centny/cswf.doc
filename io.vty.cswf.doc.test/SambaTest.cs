using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using System.Collections.Generic;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class SambaTest
    {
        [TestMethod]
        public void TestSamba()
        {
            Samba.RunDelay = 100;
            Samba.ChkDelay = 500;
            IDictionary<string, int> paths = new Dictionary<string, int>();
            paths.Add("K:\\test.txt", Samba.RW);

            Samba.AddVolume("K:", "\\\\192.168.1.26\\nfs_test", "nfs_test", "sco", paths);
            new Thread((obj) =>
            {
                Samba.LoopChecker();
            }).Start();
            Thread.Sleep(3000);
            Console.WriteLine("done...");
            Samba.AddVolume2("K:", "\\\\192.168.1.26\\nfs_test", "nfs_test", "sco", "{\"X:\\test.txt\":2}");
        }
    }
}
