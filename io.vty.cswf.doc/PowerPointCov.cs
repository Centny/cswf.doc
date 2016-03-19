using io.vty.cswf.cache;
using io.vty.cswf.log;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
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
            public PowerPoint(Application app)
            {
                this.App = app;
            }

            public void Dispose()
            {
                try
                {
                    this.App.Quit();
                }
                catch (Exception e)
                {
                    L.W(e, "PPTX quit the powerpoint application fail with error->{0}", e.Message);
                }
            }
        }
        public static CachedQueue<PowerPoint> Cached = new CachedQueue<PowerPoint>(30000, 3);
        public static PowerPoint Dequeue()
        {
            PowerPoint app;
            if (Cached.TryDequeue(out app))
            {
                return app;
            }
            else
            {
                app = new PowerPoint(new Application());
                return app;
            }
        }
        public static void Enqueue(PowerPoint app)
        {
            Cached.Enqueue(app);
        }

        public String FilterName { get; set; }
        public PowerPointCov(String src, String dst_f, int beg = 0) : base(src, dst_f, 1, 1, beg)
        {
            this.FilterName = "png";
        }

        public override void Exec()
        {
            var file_c = this.Beg;
            L.D("executing ppt2img by file({0}),destination format({1})", this.AsSrc, this.AsDstF);
            var app = Dequeue();
            Presentation doc = null;
            //int pid = 0;
            this.Cdl.add();
            try
            {
                doc = app.App.Presentations.Open(this.AsSrc, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                var total = doc.Slides.Count;
                this.Total = new int[total];
                this.Done = new int[total];
                for (var i = 1; i <= total; i++)
                {
                    Slide slide = doc.Slides[i];
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
                Enqueue(app);
            }
            this.Cdl.done();
            this.Cdl.wait();
            L.D("executing ppt2img by file({0}),destination format({1}) done with slides({2}),fail({3})", this.AsSrc, this.AsDstF, file_c, this.Fails.Count);
        }

        protected virtual int Ppt2imgProc(Slide slide, int idx, int file_c)
        {
            this.Total[idx] = 1;
            var spath = String.Format(this.AsDstF, file_c);
            if (this.ShowLog)
            {
                L.D("ppt2img parsing file({0},{1}) to {2}", this.AsSrc, file_c, spath);
            }
            slide.Export(spath, this.FilterName, 0, 0);
            var rspath = String.Format(this.DstF, file_c);
            this.Result.Count += 1;
            this.Result.Files.Add(rspath);
            this.Done[idx] = 1;
            this.OnDone();
            return 1;
        }
    }
}
