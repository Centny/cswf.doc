using io.vty.cswf.log;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using ppt = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Runtime.Serialization;
using io.vty.cswf.util;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace io.vty.cswf.doc
{
    /// <summary>
    /// the Converter to convert docx/pptx/xlsx to image
    /// </summary>
    public class Converter
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        /// <summary>
        /// the log
        /// </summary>
        private static readonly ILog L = Log.New();
        /// <summary>
        /// the process
        /// </summary>
        public class Proc
        {
            /// <summary>
            /// the exe name
            /// </summary>
            public string Name
            {
                get; set;
            }
            /// <summary>
            /// destiance format string.
            /// </summary>
            public string DstF
            {
                get; set;
            }
            /// <summary>
            /// the arguments.
            /// </summary>
            public string Args
            {
                get; set;
            }
            /// <summary>
            /// the progress url by string.format
            /// </summary>
            public string Progress
            {
                get; set;
            }
            /// <summary>
            /// exec the command ans append the result to Res
            /// </summary>
            /// <param name="res">the result</param>
            /// <param name="count">currnet count</param>
            /// <param name="spath">targe file path</param>
            /// <returns></returns>
            public int exec(Res res, int total, int count, string spath)
            {
                string data = "";
                var dst_f = string.Format(this.DstF, spath, count);
                L.D("Proc start exec <{0} {1} {2}>", this.Name, dst_f, this.Args);
                var code = Exec.exec(out data, this.Name, dst_f, this.Args);
                if (code != 0)
                {
                    L.E("Proc exec <{0} {1} {2}> fail({3}) with result->\n{4}", this.Name, dst_f, this.Args, code, data);
                    throw new Exception(string.Format("Proc exec <{0}> fail with exit code({1})", this.Name, code));
                }
                data = data.Trim();
                var lines = data.Split('\n');
                var added = 0;
                foreach (var line in lines)
                {
                    var tline = line.Trim();
                    if (tline.Length < 1)
                    {
                        continue;
                    }
                    res.Files.Add(line);
                    added += 1;
                }
                if (this.Progress != null && this.Progress.Length > 0)
                {
                    var rate = ((float)count + 1) / ((float)total);
                    L.D("Proc do progress notify to {0} with rate({1})", this.Progress, rate);
                    H.doGet(this.Progress, rate);
                }
                L.D("Proc exec <{0} {1} {2}> done with {3} file added", this.Name, dst_f, this.Args, added);
                return added;
            }
            /// <summary>
            /// the constructor by exe name and destiance format.
            /// </summary>
            /// <param name="name">exe name</param>
            /// <param name="dst_f">destiance format string</param>
            public Proc(string name, string dst_f)
            {
                this.Name = name;
                this.DstF = dst_f;
            }
        }

        /// <summary>
        /// the delegate to process the converter.
        /// </summary>
        /// <param name="res">the result</param>
        /// <param name="count">the current count</param>
        /// <param name="spath">the tart file path</param>
        /// <returns></returns>
        public delegate int OnProcess(Res res, int total, int count, string spath);
        /// <summary>
        /// the struct of result
        /// </summary>
        [DataContract]
        public class Res
        {
            /// <summary>
            /// the result code, 0 is success, other is fail.
            /// </summary>
            [DataMember(Name = "code")]
            public int Code
            {
                get; set;
            }
            /// <summary>
            /// the result count.
            /// </summary>
            [DataMember(Name = "count")]
            public int Count
            {
                get; set;
            }
            /// <summary>
            /// the result data.
            /// </summary>
            [DataMember(Name = "files")]
            public IList<string> Files
            {
                get; set;
            }
            /// <summary>
            /// the souce file
            /// </summary>
            [DataMember(Name = "src")]
            public string Src
            {
                get; set;
            }
            /// <summary>
            /// constructor by souce file path.
            /// </summary>
            /// <param name="src">the source file path</param>
            public Res(string src)
            {
                this.Src = src;
                this.Files = new List<string>();
            }
            /// <summary>
            /// saving the result to file with json format.
            /// </summary>
            /// <param name="json"></param>
            public void Save(string json)
            {
                using (var sw = new StreamWriter(json))
                {
                    sw.Write(Json.stringify(this));
                }
            }
            public void Trim(string prefix)
            {
                for (var i = 0; i < this.Files.Count; i++)
                {
                    this.Files[i] = this.Files[i].Replace(prefix, "");
                }
            }
        }
        /// <summary>
        /// execute command and convert result to Res
        /// </summary>
        /// <param name="src">the source file path</param>
        /// <param name="dst_f">the destiance output path with string format</param>
        /// <param name="beg">the begin number of format string</param>
        /// <param name="log">if show detail log</param>
        /// <param name="process">the process delegate</param>
        /// <returns>the result</returns>
        public static Res exec(string src, string dst_f, int beg = 0, bool log = false, OnProcess process = null)
        {
            var res = new Res(src);
            var count = beg;
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            L.D("executing exec by file({0}),destination format({1})", as_src, as_dst_f);
            var rspath = string.Format(dst_f, count);
            if (process == null)
            {
                res.Files.Add(rspath);
                count += 1;
            }
            else
            {
                count += process(res, 1, count, rspath);
            }
            res.Count = count;
            L.D("executing exec by file({0}),destination format({1}) done with count({2})", as_src, as_dst_f, count);
            return res;
        }


        /// <summary>
        /// convert word to png
        /// </summary>
        /// <param name="src">the source file</param>
        /// <param name="dst_f">the destinace out file path formatiing with page number, like xxx-{0}.png</param>
        /// <param name="beg">the begin number of format string</param>
        /// <param name="log">if show detail log</param>
        /// <param name="process">the process delegate</param>
        /// <returns>the result</returns>
        public static Res word2img(string src, string dst_f, int beg = 0, bool log = false, OnProcess process = null)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var pages = beg;
            L.D("executing word2png by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new word.Application();
            Process proc = null;
            try
            {
                app.Visible = true;
                var doc = app.Documents.Open(as_src, false, true);
                doc.ShowGrammaticalErrors = false;
                doc.PrintFormsData = false;
                //doc.ShowRevisions = false;
                doc.ShowSpellingErrors = false;
                if (doc.Windows.Count < 1)
                {
                    L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", as_src, as_dst_f, pages);
                    doc.Close(false);
                    res.Code = 404;
                    return res;
                }
                try
                {
                    int procid = 0;
                    GetWindowThreadProcessId(app.ActiveWindow.Hwnd, out procid);
                    proc = Process.GetProcessById(procid);
                }
                catch (Exception e)
                {
                    L.W(e, "executing word2png by file({0}),destination format({1}) get process error->{2}", as_src, as_dst_f, e.Message);
                }
                word.Window window = doc.Windows[1];
                //foreach (word.Window window in doc.Windows)
                //{
                if (window.Panes.Count < 1)
                {
                    L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", as_src, as_dst_f, pages);
                    doc.Close(false);
                    res.Code = 404;
                    return res;
                }
                word.Pane pane = window.Panes[1];
                //foreach (word.Pane pane in window.Panes)
                //{
                if (log)
                {
                    L.D("executing word2png by file({0}),destination format({1}) with {2} page found", as_src, as_dst_f, pane.Pages.Count);
                }
                var total = pane.Pages.Count;
                for (var i = 1; i <= total; i++)
                {
                    var spath = String.Format(as_dst_f, pages);
                    if (log)
                    {
                        L.D("word2png parsing file({0},{1}) to {2}", as_src, pages, spath);

                    }
                    var page = pane.Pages[i];
                    dynamic bits;
                    try
                    {
                        bits = page.EnhMetaFileBits;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                    using (var ms = new MemoryStream((byte[])(bits)))
                    {
                        Image.FromStream(ms).Save(spath, ImageFormat.Png);
                    }
                    var rspath = String.Format(dst_f, pages);
                    if (process == null)
                    {
                        res.Files.Add(rspath);
                        pages += 1;
                    }
                    else
                    {
                        pages += process(res, total, pages, rspath);
                    }

                }
                //  panes += 1;
                //}
                //}
                L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", as_src, as_dst_f, pages);
                res.Code = 0;
                res.Count = pages;
                
                for (int i = 0; i < 100; i++)
                {
                    try
                    {
                        doc.Close(false);
                        break;
                    }
                    catch (Exception e)
                    {
                        L.W(e, "executing word2png by file({0}),destination format({1}) closing error->{2}", as_src, as_dst_f, e.Message);
                        SetForegroundWindow(new IntPtr(app.ActiveWindow.Hwnd));
                        app.Activate();
                        SendKeys.SendWait("\n\n");
                    }
                }
                app.Quit(false);
            }
            catch (Exception e)
            {
                L.E(e, "executing word2png by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f, e.Message);
                res.Code = 500;
                app.Quit(false);
                throw e;
            }
            finally
            {
                try
                {
                    while (proc != null && !proc.HasExited)
                    {
                        Thread.Sleep(1000);
                        proc.Kill();
                    }
                }
                catch (Exception e)
                {
                    L.W(e, "executing word2png by file({0}),destination format({1}) kill process error->{2}", as_src, as_dst_f, e.Message);
                }
            }
            return res;
        }

        /// <summary>
        /// convert exec to pdf
        /// </summary>
        /// <param name="src">the source file</param>
        /// <param name="dst_f">the destinace out file path formatiing with page number, like xxx-{0}.pdf</param>
        /// <param name="beg">the begin number of format string</param>
        /// <param name="log">if show detail log</param>
        /// <param name="process">the process delegate</param>
        /// <returns>the result</returns>
        public static Res excel2pdf(String src, String dst_f, int beg = 0, bool log = false, OnProcess process = null)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var sheets = beg;
            L.D("executing excel2pdf by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new excel.Application();
            Process proc = null;
            try
            {
                app.Visible = true;
                var books = app.Workbooks.Open(as_src, 0, true, 5, "", "",
                    true, excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                var total = books.Worksheets.Count;
                try {
                    int procid = 0;
                    GetWindowThreadProcessId(app.ActiveWindow.Hwnd, out procid);
                    proc = Process.GetProcessById(procid);
                }catch(Exception e)
                {
                    L.W(e, "executing excel2pdf by file({0}),destination format({1}) get process error->{2}", as_src, as_dst_f, e.Message);
                }
                foreach (excel.Worksheet sheet in books.Worksheets)
                {
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
                    var spath = String.Format(as_dst_f, sheets);
                    if (log)
                    {
                        L.D("excel2pdf parsing file({0},{1}) to {2}", as_src, sheets, spath);
                    }
                    sheet.ExportAsFixedFormat(excel.XlFixedFormatType.xlTypePDF, spath);
                    var rspath = String.Format(dst_f, sheets);
                    if (process == null)
                    {
                        res.Files.Add(rspath);
                        sheets += 1;
                    }
                    else
                    {
                        sheets += process(res, total, sheets, rspath);
                    }
                }
                L.D("executing excel2pdf by file({0}),destination format({1}) done with sheets({2})", as_src, as_dst_f, sheets);
                res.Code = 0;
                res.Count = sheets;
                for (int i = 0; i < 100; i++)
                {
                    try
                    {
                        books.Close(false, null, null);
                        break;
                    }
                    catch (Exception e)
                    {
                        L.E(e, "executing excel2pdf by file({0}),destination format({1}) closing error->{2}", as_src, as_dst_f, e.Message);
                        SetForegroundWindow(new IntPtr(app.ActiveWindow.Hwnd));
                        app.ActiveWindow.Activate();
                        SendKeys.SendWait("\n\n");
                    }
                }
                app.Quit();
            }
            catch (Exception e)
            {
                L.E(e, "executing excel2pdf by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f, e.Message);
                res.Code = 500;
                throw e;
            }
            finally
            {
                try
                {
                    while (proc != null && !proc.HasExited)
                    {
                        Thread.Sleep(1000);
                        proc.Kill();
                    }
                }
                catch (Exception e)
                {
                    L.W(e, "executing excel2pdf by file({0}),destination format({1}) kill process error->{2}", as_src, as_dst_f, e.Message);
                }
            }
            return res;
        }
        /// <summary>
        /// convert ppt to image
        /// </summary>
        /// <param name="src">the source file</param>
        /// <param name="dst_f">the destinace out file path formatiing with page number, like xxx-{0}.png</param>
        /// <param name="beg">the begin number of format string</param>
        /// <param name="filterName">the file name to image format</param>
        /// <param name="scaleWidth">scale with</param>
        /// <param name="scaleHeight">scale height</param>
        /// <param name="log">if show detail log</param>
        /// <param name="process">the process delegate</param>
        /// <returns>the result</returns>
        public static Res ppt2img(String src, String dst_f, int beg = 0, string filterName = "png", int scaleWidth = 0, int scaleHeight = 0, bool log = false, OnProcess process = null)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var slides = beg;
            L.D("executing ppt2img by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new ppt.Application();
            Process proc = null;
            try
            {
                var doc = app.Presentations.Open(as_src, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                try
                {
                    int procid = 0;
                    GetWindowThreadProcessId(app.HWND, out procid);
                    proc = Process.GetProcessById(procid);
                }
                catch (Exception e)
                {
                    L.W(e, "executing ppt2img by file({0}),destination format({1}) get process error->{2}", as_src, as_dst_f, e.Message);
                }
                var total = doc.Slides.Count;
                foreach (ppt.Slide slide in doc.Slides)
                {
                    var spath = String.Format(as_dst_f, slides);
                    if (log)
                    {
                        L.D("ppt2img parsing file({0},{1}) to {2}", as_src, slides, spath);
                    }
                    slide.Export(spath, filterName, scaleWidth, scaleHeight);
                    var rspath = String.Format(dst_f, slides);
                    if (process == null)
                    {
                        res.Files.Add(rspath);
                        slides += 1;
                    }
                    else
                    {
                        slides += process(res, total, slides, rspath);
                    }

                }
                L.D("executing ppt2img by file({0}),destination format({1}) done with slides({2})", as_src, as_dst_f, slides);
                res.Code = 0;
                res.Count = slides;
                app.Quit();
            }
            catch (Exception e)
            {
                L.E(e, "executing ppt2img by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f, e.Message);
                res.Code = 500;
                app.Quit();
                throw e;
            }
            finally
            {
                try
                {
                    while (proc != null && !proc.HasExited)
                    {
                        Thread.Sleep(1000);
                        proc.Kill();
                    }
                }
                catch (Exception e)
                {
                    L.W(e, "executing ppt2img by file({0}),destination format({1}) kill process error->{2}", as_src, as_dst_f, e.Message);
                }
            }
            return res;
        }
    }
}
