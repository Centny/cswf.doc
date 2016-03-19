using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace io.vty.cswf.doc.test
{
    [TestClass]
    public class ErrTest
    {
        public class ErrProc : CovProc
        {
            public ErrProc(int a, int b) : base("xx", "xx", a, b, 0)
            {

            }

            public override void Exec()
            {
                throw new NotImplementedException();
            }
            public void test_done()
            {
                base.OnDone();
                this.Total = new int[] { };
                base.OnDone();
                this.Done = new int[] { };
                base.OnDone();
                this.Total = new int[] { 1 };
                base.OnDone();
                this.Done = new int[] { 0 };
                base.OnDone();
            }
        }
        [TestMethod]
        public void TestCovProc()
        {
            try
            {
                new ErrProc(0, 111);
                Assert.Fail();
            }
            catch (Exception)
            {

            }
            try
            {
                new ErrProc(11, 0);
                Assert.Fail();
            }
            catch (Exception)
            {

            }
            var cov = new ErrProc(11, 11);
            cov.test_done();
        }
    }
}
