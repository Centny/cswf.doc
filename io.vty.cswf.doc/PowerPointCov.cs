using io.vty.cswf.cache;
using io.vty.cswf.log;
using io.vty.cswf.util;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class PowerPointCov : CovProc
    {
        private static readonly ILog L = Log.New();
        public class PowerPoint : IDisposable
        {
            public Application App;
            public Presentation Doc;
            public int Pid;
            public PowerPoint(Application app)
            {
                this.App = app;
            }

            public void Dispose()
            {
                try
                {
                    this.App.Quit();
                    //ProcKiller.DelRunning(this.Pid);
                    L.D("PowerPoint application({0}) quit success", this.Pid);
                }
                catch (Exception e)
                {
                    L.W(e, "PPTX quit the powerpoint application fail with error->{0}", e.Message);
                }
            }
        }
        public static CachedQueue<PowerPoint> Cached = new CachedQueue<PowerPoint>(30000, 3);
        public static PowerPoint Dequeue(string src)
        {
            PowerPoint app;
            if (Cached.TryDequeue(out app))
            {
                try
                {
                    app.Doc = app.App.Presentations.Open(src, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                }
                catch (Exception e)
                {
                    Cached.Enqueue(app);
                    throw e;
                }
                return app;
            }
            try
            {
                //ProcKiller.Shared.Lock();
                app = new PowerPoint(new Application());
                //app.App.Visible = MsoTriState.msoTrue;
                app.Doc = app.App.Presentations.Open(src, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                app.Pid = CovProc.GetWindowThreadProcessId(app.App.HWND);
                ProcKiller.MarkUsed(app.Pid);
                //ProcKiller.AddRunning(app.Pid);
                //ProcKiller.Shared.Unlock();
            }
            catch (Exception e)
            {
                //ProcKiller.Shared.Unlock();
                throw e;
            }
            return app;
        }
        public static void Enqueue(PowerPoint app)
        {
            try
            {
                if (app.Doc != null)
                {
                    app.Doc.Close();
                }
                app.Doc = null;
                Cached.Enqueue(app);
                ProcKiller.MarkDone(app.Pid);
            }
            catch (Exception e)
            {
                L.W(e, "Close PowerPoint fail with error->", e.Message);
            }

        }

        public String FilterName { get; set; }
        public PowerPointCov(String src, String dst_f, int beg = 0) : base(src, dst_f, 1, 1, beg)
        {
            this.FilterName = "png";
        }

        public override void Exec()
        {
            var file_c = 0;
            L.D("executing ppt2img by file({0}),destination format({1})", this.AsSrc, this.AsDstF);
            PowerPoint app = null;
            this.Cdl.add();
            var tf_base = Path.GetTempFileName();
            var tf = tf_base + Path.GetExtension(this.AsSrc);
            try
            {
                File.Copy(this.AsSrc, tf, true);
                app = Dequeue(tf);
                var total = app.Doc.Slides.Count;
                if (total > this.MaxPage)
                {
                    this.Result.Code = 413;
                    return;
                }
                this.Total = new int[total];
                this.Done = new int[total];
                Util.set(this.Total, 1);
                Util.set(this.Done, 0);
                for (var i = 1; i <= total; i++)
                {
                    Slide slide = app.Doc.Slides[i];
                    file_c += this.Ppt2imgProc(slide, i - 1, file_c);
                }
            }
            catch (Exception e)
            {
                L.E(e, "executing ppt2img by file({0}),destination format({1}) done with error->{2}", this.AsSrc, this.AsDstF, e.Message);
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
            finally
            {
                if (app != null)
                {
                    Enqueue(app);
                }
                try
                {
                    File.Delete(tf);
                    File.Delete(tf_base);
                }
                catch (Exception e)
                {
                    L.W("executing ppt2img on delete temp file({0}) error->{1}", tf, e.Message);
                }
            }
            this.Cdl.done();
            this.Cdl.wait();
            L.D("executing ppt2img by file({0}),destination format({1}) done with slides({2}),fail({3})", this.AsSrc, this.AsDstF, file_c, this.Fails.Count);
        }

        protected virtual int Ppt2imgProc(Slide slide, int idx, int file_c)
        {
            this.Total[idx] = 1;
            var spath = String.Format(this.AsDstF, file_c);
            var as_dir = Path.GetDirectoryName(spath);
            if (!Directory.Exists(as_dir))
            {
                Directory.CreateDirectory(as_dir);
            }
            if (this.ShowLog)
            {
                L.D("ppt2img parsing file({0},{1}) to {2}", this.AsSrc, file_c, spath);
            }
            slide.Export(spath, this.FilterName, 0, 0);
            var rspath = String.Format(this.DstF, this.Beg + file_c);
            this.Result.Count += 1;
            this.Result.Files.Add(rspath);
            this.Done[idx] = 1;
            this.OnDone();
            return 1;
        }
    }
}
