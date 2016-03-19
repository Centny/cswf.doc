using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public abstract class CovProc : IDisposable
    {
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        public static int GetWindowThreadProcessId(int hwnd)
        {
            int pid = 0;
            GetWindowThreadProcessId(hwnd, out pid);
            return pid;
        }
        public delegate void OnProc(CovProc cov, float rate);
        public bool ShowLog = false;
        public String Src { get; protected set; }
        public String DstF { get; protected set; }
        public int Beg { get; protected set; }
        protected String AsSrc;
        protected String AsDstF;
        public CDL Cdl { get; protected set; }
        public int[] Total { get; protected set; }
        public int[] Done { get; protected set; }
        public CovRes Result { get; protected set; }
        public int MaxWidth { get; protected set; }
        public int MaxHeight { get; protected set; }
        public IList<Exception> Fails { get; protected set; }
        public long LastProc { get; protected set; }
        public long ProcDelay { get; set; }
        public OnProc Proc { get; set; }
        public object State;
        public CovProc(String src, String dst_f, int maxw = 768, int maxh = 1024, int beg = 0)
        {
            this.Src = src;
            this.DstF = dst_f;
            this.AsSrc = Path.GetFullPath(src);
            this.AsDstF = Path.GetFullPath(dst_f);
            this.Result = new CovRes(this.Src);
            this.Cdl = new CDL(0);
            this.Fails = new List<Exception>();
            this.MaxWidth = maxw;
            this.MaxHeight = maxh;
            this.Beg = beg;
            this.ProcDelay = 1000;
            if (this.MaxWidth < 1 || this.MaxHeight < 1)
            {
                throw new ArgumentException("the maxw/maxh must be greater zero");
            }
        }

        public abstract void Exec();
        public virtual void Dispose()
        {

        }

        public virtual void PrintFails()
        {
            Console.WriteLine(this);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var e in this.Fails)
            {
                sb.Append("Fails->" + e.Message + "\n");
                sb.Append(e.StackTrace + "\n");
                sb.Append("------->\n");
            }
            return sb.ToString();
        }

        protected void OnDone()
        {
            var now = Util.Now();
            if (now - this.LastProc < this.ProcDelay)
            {
                return;
            }
            if (this.Total == null || this.Done == null || this.Total.Length < 1 || this.Total.Length != this.Done.Length || this.Proc == null)
            {
                return;
            }
            float rate = 0;
            for (var i = 0; i < this.Total.Length; i++)
            {
                rate += ((float)this.Done[i]) / ((float)this.Total[i]);
            }
            rate = rate / ((float)this.Total.Length);
            this.Proc(this, rate);
            this.LastProc = now;
        }
    }
}
