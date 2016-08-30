using io.vty.cswf.log;
using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class WindowCloser
    {
        public static WindowCloser Shared = new WindowCloser();
        public static void StartWindowCloser(String inc, String exc, int period = 3000)
        {
            if (!String.IsNullOrWhiteSpace(inc))
            {
                foreach (String i in inc.Split(','))
                {
                    Shared.Inc.Add(i);
                }
            }
            if (!String.IsNullOrWhiteSpace(exc))
            {
                foreach (String e in exc.Split(','))
                {
                    Shared.Exc.Add(e);
                }
            }
            Shared.Period = period;
            L.D("WindowCloser start by current names:\n{0}", Util.tos(Shared.ListCurrent()));
            Shared.Start();
        }
        public static void StopWindowCloser()
        {
            Shared.Stop();
        }

        public static readonly ILog L = Log.New();
        public const int WM_CLOSE = 0x10;
        public delegate bool EnumWindowsProc(int hWnd, int param);
        [DllImport("user32.dll")]
        private static extern int EnumWindows(EnumWindowsProc proc, int param);
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, int wMsg, int wParam, int iParam);
        [DllImport("user32.dll")]
        public static extern int GetWindowText(int hWnd, StringBuilder title, int size);
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(int hWnd);
        public IList<String> Exc { protected set; get; }
        public IList<String> Inc { protected set; get; }
        public Timer T { get; protected set; }
        public int Period { get; set; }
        public WindowCloser()
        {
            this.Exc = new List<String>();
            this.Inc = new List<String>();
        }

        public void SendClose()
        {
            EnumWindows(this.doProc, 0);
        }

        public void ExcCurrent()
        {
            EnumWindows((hWnd, param) =>
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }
                StringBuilder title = new StringBuilder(10240);
                GetWindowText(hWnd, title, title.Capacity);
                String msg = title.ToString();
                if (String.IsNullOrWhiteSpace(msg))
                {
                    return true;
                }
                this.Exc.Add(msg);
                return true;
            }, 0);
        }
        public IList<String> ListCurrent()
        {
            var names = new List<String>();
            EnumWindows((hWnd, param) =>
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }
                StringBuilder title = new StringBuilder(10240);
                GetWindowText(hWnd, title, title.Capacity);
                String msg = title.ToString();
                if (String.IsNullOrWhiteSpace(msg))
                {
                    return true;
                }
                names.Add(msg);
                return true;
            }, 0);
            return names;
        }
        protected virtual bool doProc(int hWnd, int param)
        {
            if (!IsWindowVisible(hWnd))
            {
                return true;
            }
            StringBuilder title = new StringBuilder(10240);
            try
            {
                GetWindowText(hWnd, title, title.Capacity);
                String msg = title.ToString();
                if (String.IsNullOrWhiteSpace(msg))
                {
                    return true;
                }
                if (!this.isHitted(msg))
                {
                    L.D("WindowCloser do proc ignore widown({0})", title.ToString());
                    return true;
                }
                SendMessage(hWnd, WM_CLOSE, 0, 0);
                L.D("doProc sending close message to window({0}) success", title.ToString());
            }
            catch (Exception e)
            {
                L.E("doProc for window({0}) fail with error({1})", title.ToString(), e.Message, e);
            }
            return true;
        }
        protected virtual bool isHitted(String title)
        {
            foreach (String exc in this.Exc)
            {
                if (title.IndexOf(exc, StringComparison.OrdinalIgnoreCase) > -1)
                {
                    return false;
                }
            }
            foreach (String inc in this.Inc)
            {
                if (title.IndexOf(inc, StringComparison.OrdinalIgnoreCase) > -1)
                {
                    return true;
                }
            }
            return false;
        }
        public void Start()
        {
            this.T = new Timer((o) =>
            {
                this.SendClose();
            }, 0, this.Period, this.Period);
        }

        public void Stop()
        {
            this.Dispose();
        }
        public void Dispose()
        {
            if (this.T == null)
            {
                return;
            }
            this.T.Dispose();
            this.T = null;
        }
    }
}
