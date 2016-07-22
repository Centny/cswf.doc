using io.vty.cswf.netw;
using io.vty.cswf.netw.dtm;
using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using io.vty.cswf.netw.impl;
using io.vty.cswf.log;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace io.vty.cswf.doc
{
    public class DocCov : DTM_C_j
    {
        [DllImport("user32.dll")]
        static extern bool TerminateProcess(IntPtr hwnd, uint code);

        public static void CloseProc(Process proc)
        {
            try {
                TerminateProcess(proc.Handle, 1);
            }catch(Exception)
            {
            }
        }
        public enum SupportedL
        {
            None = 0,
            Word = 1,
            Excel = 2,
            PowerPoint = 3,
            Pdf = 4,
            Img = 5,
        }
        public static SupportedL parseSupported(String key)
        {
            if ("Word".Equals(key))
            {
                return SupportedL.Word;
            }
            /*
            else if ("Excel".Equals(key))
            {
                return SupportedL.Excel;
            }
            */
            else if ("PowerPoint".Equals(key))
            {
                return SupportedL.PowerPoint;
            }
            else if ("Pdf".Equals(key))
            {
                return SupportedL.Pdf;
            }
            else if ("Img".Equals(key))
            {
                return SupportedL.Img;
            }
            else
            {
                return SupportedL.None;
            }
        }
        private static readonly ILog L = Log.New();
        public DocCov(string name, FCfg cfg) : base(name, cfg)
        {
            this.builder = () =>
            {
                throw new NotImplementedException("NetwBaseBuilder is not implemented");
            };
        }
        public DocCov(string name, FCfg cfg, NetwRunnerV.NetwBaseBuilder builder) : base(name, cfg, builder)
        {
            this.builder = builder;
        }

        public void InitConfig()
        {
            WordCov.Cached.MaxIdle = this.Cfg.Val("word_idle", 5);
            ExcelCov.Cached.MaxIdle = this.Cfg.Val("excel_idle", 5);
            PowerPointCov.Cached.MaxIdle = this.Cfg.Val("power_point_idle", 5);
            TaskPool.Shared.MaximumConcurrency = this.Cfg.Val("max_tasks", 32);
            ThreadPool.SetMaxThreads(this.Cfg.Val("max_worker_threads", 16), this.Cfg.Val("max_async_threads", 16));
            ThreadPool.SetMaxThreads(this.Cfg.Val("min_worker_threads", 4), this.Cfg.Val("min_async_threads", 4));
        }
        public void StartMonitor()
        {
            var names = this.Cfg.Val("MPNS", "");
            if (names.Length < 1)
            {
                return;
            }
            ProcKiller.Shared.OnClose = CloseProc;
            var period = this.Cfg.Val("MPPT", 30000);
            L.I("DocCov start process monitor by names({0}),period({1})", names, period);
            foreach (var name in names.Split(','))
            {
                ProcKiller.AddName(name);
            }
            ProcKiller.StartTimer(period);
        }
        public override void DoCmd(string tid, FCfg fcfg, string cmds)
        {
            var args = Args.parseArgs(cmds);
            String cmd;
            if (!args.StringVal(0, out cmd))
            {
                throw new ArgumentException("DocCov receive emtpy command", "cmds");
            }
            SupportedL sp = parseSupported(cmd);
            if (SupportedL.None.Equals(sp))
            {
                base.DoCmd(tid, fcfg, cmds);
            }
            else
            {
                this.RunSupported(tid, sp, fcfg, args, cmds);
            }
        }
        protected virtual void RunSupported(String tid, SupportedL sp, FCfg cfg, Args args, String cmds)
        {
            String src, dst_f;
            if (!(args.StringVal(1, out src) && args.StringVal(2, out dst_f)))
            {
                throw new ArgumentException("Word argument is invalid, please confirm arguments using by <src dst_f maxw maxh>");
            }
            ThreadPool.QueueUserWorkItem(this.RunSupportedProc_, new object[] { tid, sp, cfg, args, cmds, src, dst_f });
        }
        private void RunSupportedProc_(object state)
        {
            var args = state as object[];
            this.RunSupportedProc((String)args[0], (SupportedL)args[1], (FCfg)args[2], (Args)args[3], (String)args[4], (String)args[5], (String)args[6]);
        }
        protected virtual void RunSupportedProc(String tid, SupportedL sp, FCfg cfg, Args args, String cmds, String src, String dst_f)
        {
            L.I("DocCov calling Supported({2}) by (\n{0}\n) by tid({1})", cmds, tid, sp);
            var beg = Util.Now();
            var rargs = Util.NewDict();
            rargs["tid"] = tid;
            try
            {
                CovProc cov = this.RunSupported(tid, sp, cfg, args, src, dst_f);
                rargs["code"] = cov.Result.Code;
                if (cov.Fails.Count > 0)
                {
                    rargs["err"] = String.Format("{0} exeception found, see DocCov log for detail", cov.Fails.Count);
                    L.E("DocCov calling Supported({3}) by (\n{0}\n) by tid({1}) fail with->\n{2}\n", cmds, tid, cov.ToString(), sp);
                }
                else
                {
                    rargs["data"] = cov.Result;
                }
            }
            catch (Exception e)
            {
                rargs["code"] = 500;
                rargs["err"] = String.Format("{0} exeception found, see DocCov log for detail", 1);
                L.E(e, "DocCov calling Supported({3}) by (\n{0}\n) by tid({1}) fail with error->{2}", cmds, tid, e.Message, sp);
            }
            var used = Util.Now() - beg;
            rargs["used"] = used;
            try
            {
                this.SendDone(rargs);
                L.I("DocCov calling Supported({2}) success by (\n{0}\n) by tid({1})", cmds, tid, sp);
            }
            catch (Exception e)
            {
                L.E(e, "DocCov calling Supported({3}) by (\n{0}\n) by tid({1}) fail with send done err->", cmds, tid, e.Message, sp);
            }
        }
        protected virtual CovProc RunSupported(String tid, SupportedL sp, FCfg cfg, Args args, String src, String dst_f)
        {
            CovProc cov = null;
            String prefix = "";
            int maxw, maxh;
            switch (sp)
            {
                case SupportedL.Word:
                    args.IntVal(3, out maxw, 768);
                    args.IntVal(4, out maxh, 1024);
                    args.StringVal(5, out prefix, "");
                    cov = new WordCov(src, dst_f, maxw, maxh);
                    break;
                case SupportedL.Excel:
                    args.IntVal(3, out maxw, 768);
                    args.IntVal(4, out maxh, 1024);
                    args.StringVal(5, out prefix, "");
                    cov = new ExcelCov(src, dst_f, maxw, maxh, cfg.Val("density_x", 96), cfg.Val("density_y", 96));
                    break;
                case SupportedL.PowerPoint:
                    args.StringVal(3, out prefix, "");
                    cov = new PowerPointCov(src, dst_f);
                    break;
                case SupportedL.Pdf:
                    args.IntVal(3, out maxw, 768);
                    args.IntVal(4, out maxh, 1024);
                    args.StringVal(5, out prefix, "");
                    cov = new PdfCov(src, dst_f, maxw, maxh, cfg.Val("density_x", 96), cfg.Val("density_y", 96));
                    break;
                case SupportedL.Img:
                    args.IntVal(3, out maxw, 768);
                    args.IntVal(4, out maxh, 1024);
                    args.StringVal(5, out prefix, "");
                    cov = new ImgCov(src, dst_f, maxw, maxh);
                    break;
                default:
                    throw new ArgumentException("the not supported command", "sp");
            }
            cov.State = tid;
            cov.Proc = this.OnCovProc;
            cov.ShowLog = cfg.Val("showlog", 0) == 1;
            cov.Exec();
            cov.Dispose();
            if (prefix.Length > 0)
            {
                cov.Result.Trim(prefix);
            }
            return cov;
        }

        protected virtual void OnCovProc(CovProc cov, float rate)
        {
            this.NotifyProc(cov.State as string, rate);
        }
    }
}
