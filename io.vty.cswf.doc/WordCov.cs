using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using io.vty.cswf.log;
using io.vty.cswf.cache;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using io.vty.cswf.util;
using System.Diagnostics;

namespace io.vty.cswf.doc
{
    public class WordCov : CovProc
    {
        private static readonly ILog L = Log.New();
        public class Word : IDisposable
        {
            public Application App;
            public Document Doc;
            public int Pid;
            public Word(Application app)
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
        public static CachedQueue<Word> Cached = new CachedQueue<Word>(30000, 3);
        public static Word Dequeue(string src)
        {
            Word app;
            if (Cached.TryDequeue(out app))
            {
                app.Doc = app.App.Documents.Open(src, false, true);
                app.Doc.ShowGrammaticalErrors = false;
                app.Doc.PrintFormsData = false;
                app.Doc.ShowSpellingErrors = false;
            }
            else
            {
                try
                {
                    ProcKiller.Shared.Lock();
                    app = new Word(new Application());
                    app.App.Visible = true;
                    app.Doc = app.App.Documents.Open(src, false, true);
                    app.Doc.ShowGrammaticalErrors = false;
                    app.Doc.PrintFormsData = false;
                    app.Doc.ShowSpellingErrors = false;
                    app.Pid = CovProc.GetWindowThreadProcessId(app.Doc.ActiveWindow.Hwnd);
                    ProcKiller.AddRunning(app.Pid);
                }
                finally
                {
                    ProcKiller.Shared.Unlock();
                }
            }
            return app;
        }
        public static void Enqueue(Word app)
        {
            try
            {
                if (app.Doc != null)
                {
                    app.Doc.Close(false);
                }
                app.Doc = null;
                Cached.Enqueue(app);
            }
            catch (Exception e)
            {
                L.W(e, "Close Word document fail with error->", e.Message);
                app.Dispose();
            }

        }

        //public long LastProc;
        public WordCov(String src, String dst_f, int maxw = 768, int maxh = 1024, int beg = 0) : base(src, dst_f, maxw, maxh, beg)
        {
        }

        public override void Exec()
        {
            var pages = this.Beg;
            L.D("executing word2png by file({0}),destination format({1})", this.AsSrc, this.AsDstF);
            Word word = null;
            this.Cdl.add();
            try
            {
                word = Dequeue(this.AsSrc);
                if (word.Doc.Windows.Count < 1)
                {
                    this.Result.Code = 404;
                    L.D("executing word2png by file({0}),destination format({1}) done with code({2}),count({3})",
                        this.AsSrc, this.AsDstF, this.Result.Code, this.Result.Count);
                    return;
                }
                Window window = word.Doc.Windows[1];
                if (window.Panes.Count < 1)
                {
                    this.Result.Code = 404;
                    L.D("executing word2png by file({0}),destination format({1}) done with code({2}),count({3})",
                        this.AsSrc, this.AsDstF, this.Result.Code, this.Result.Count);
                    return;
                }
                Pane pane = window.Panes[1];
                this.Total = new int[pane.Pages.Count];
                this.Done = new int[pane.Pages.Count];
                L.D("executing word2png by file({0}),destination format({1}) with {2} page found", this.AsSrc, this.AsDstF, pane.Pages.Count);
                for (var i = 1; i <= pane.Pages.Count; i++)
                {
                    var page = pane.Pages[i];
                    byte[] bits;
                    try
                    {
                        bits = page.EnhMetaFileBits;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                    pages += this.Word2imgProc(bits, i - 1, pages);

                }
            }
            catch (Exception e)
            {
                L.E(e, "executing word2png by file({0}),destination format({1}) done with error->{2}", this.AsSrc, this.AsDstF, e.Message);
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
            finally
            {
                if (word != null)
                {
                    Enqueue(word);
                }
            }
            this.Cdl.done();
            this.Cdl.wait();
            L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", this.AsSrc, this.AsDstF, this.Result.Count);
        }
        protected virtual int Word2imgProc(byte[] bits, int idx, int pages)
        {
            this.Total[idx] = 1;
            this.Cdl.add();
            TaskPool.Queue(i =>
            {
                this.RunWord2imgProc(bits, idx, pages);
            }, idx);
            return 1;
        }

        protected void RunWord2imgProc(byte[] bits, int tidx, int pages)
        {
            if (this.Fails.Count > 0)
            {
                this.Cdl.done();
                return;
            }
            try
            {
                var as_dst = String.Format(this.AsDstF, pages);
                if (this.ShowLog)
                {
                    L.D("word2img parsing file({0},{1}) to {2}", this.AsSrc, pages, as_dst);
                }
                this.Total[tidx] = 1;
                var buf = new MemoryStream(bits);
                var img = new Bitmap(buf);
                buf.Dispose();
                Util.SaveThumbnail(img, as_dst, this.MaxWidth, this.MaxHeight, true, true, ".JPG");
                this.Result.Count += 1;
                this.Result.Files.Add(String.Format(this.DstF, pages));
                this.Done[tidx] = 1;
                this.OnDone();
            }
            catch (Exception e)
            {
                this.Fails.Add(e);
            }
            this.Cdl.done();
        }
    }
}
