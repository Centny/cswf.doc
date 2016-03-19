using ImageMagick;
using io.vty.cswf.cache;
using io.vty.cswf.log;
using io.vty.cswf.util;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class ExcelCov : PdfCov
    {

        private static readonly ILog L = Log.New();
        public class Excel : IDisposable
        {
            public Application App;
            public Workbook Book;
            public int Pid;
            public Excel(Application app)
            {
                this.App = app;
            }

            public void Dispose()
            {
                try
                {
                    this.App.Quit();
                    ProcKiller.DelRunning(this.Pid);
                }
                catch (Exception e)
                {
                    L.W(e, "Word quit the word application fail with error->{0}", e.Message);
                }
            }
        }
        public static CachedQueue<Excel> Cached = new CachedQueue<Excel>(30000, 3);
        public static Excel Dequeue(string src)
        {
            Excel app;
            if (Cached.TryDequeue(out app))
            {
                app.Book = app.App.Workbooks.Open(src, 0, true, 5, "", "",
                    true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            }
            else
            {
                try
                {
                    ProcKiller.Shared.Lock();
                    app = new Excel(new Application());
                    app.App.Visible = true;
                    app.Book = app.App.Workbooks.Open(src, 0, true, 5, "", "",
                        true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    app.Pid = CovProc.GetWindowThreadProcessId(app.App.ActiveWindow.Hwnd);
                    ProcKiller.AddRunning(app.Pid);
                }
                finally
                {
                    ProcKiller.Shared.Unlock();
                }
            }
            return app;
        }
        public static void Enqueue(Excel app)
        {
            try
            {
                app.Book.Close(false, null, null);
                Cached.Enqueue(app);
            }
            catch (Exception e)
            {
                L.E(e, "Close Excel fail with error->", e.Message);
                app.Dispose();
            }
        }

        private ICollection<MagickImageCollection> Images = new List<MagickImageCollection>();
        public ExcelCov(String src, String dst_f, int maxw = 768, int maxh = 1024, double densityx = 96, double densityy = 96, int beg = 0) : base(src, dst_f, maxw, maxh, densityx, densityy, beg)
        {


        }
        public override void Exec()
        {
            var file_c = this.Beg;
            L.D("executing excel2pdf by file({0}),destination format({1})", this.AsSrc, this.AsDstF);
            Excel app = null;
            this.Cdl.add();
            try
            {
                app = Dequeue(this.AsSrc);
                var total = app.Book.Worksheets.Count;
                this.Total = new int[total];
                this.Done = new int[total];
                for (var i = 1; i <= total; i++)
                {
                    Worksheet sheet = app.Book.Worksheets[i];
                    var range = sheet.UsedRange;
                    int rows = range.Rows.Count;
                    int cols = range.Columns.Count;
                    if (rows < 2 && cols < 2)
                    {
                        Object text = range.Text;
                        if (text is string && ((string)text).Length < 1)
                        {
                            continue;
                        }
                    }
                    file_c += this.Excel2pdfProc(sheet, i - 1, file_c);
                }

            }
            catch (Exception e)
            {
                L.E(e, "executing excel2pdf by file({0}),destination format({1}) fail with error->{2}", this.AsSrc, this.AsDstF, e.Message);
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
            finally
            {
                if (app != null)
                {
                    Enqueue(app);
                }
            }
            this.Cdl.done();
            this.Cdl.wait();
            L.D("executing excel2pdf by file({0}),destination format({1}) done with {2} image files created", this.AsSrc, this.AsDstF, this.Result.Count);
        }

        protected virtual int Excel2pdfProc(Worksheet sheet, int idx, int file_c)
        {
            if (this.Fails.Count > 0)
            {
                return 0;
            }
            var pdf = Path.GetTempFileName() + ".pdf";
            try
            {
                if (this.ShowLog)
                {
                    L.D("excel2pdf parsing file({0},{1}) to {1}", this.AsSrc, idx, pdf);
                }
                sheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdf);
                var images = new MagickImageCollection();
                this.Images.Add(images);
                var added = this.Pdf2imgProc(images, pdf, idx, file_c);
                return added;
            }
            catch (Exception e)
            {
                this.Result.Code = 500;
                this.Fails.Add(e);
                return 0;
            }
            finally
            {
                try
                {
                    File.Delete(pdf);
                }
                catch (Exception)
                {

                }
            }
        }
    }
}
