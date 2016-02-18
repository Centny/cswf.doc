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

namespace io.vty.cswf.doc
{
    public class Converter
    {
        private static readonly ILog L = Log.New();

        [DataContract]
        public class Res
        {
            [DataMember(Name = "code")]
            public int Code {
                get; set;
            }
            [DataMember(Name = "count")]
            public int Count
            {
                get; set;
            }
            [DataMember(Name = "files")]
            public IList<string> Files
            {
                get;set;
            }
            [DataMember(Name = "src")]
            public string Src
            {
                get; set;
            }
            public Res(string src)
            {
                this.Src = src;
                this.Files = new List<string>();
            }
            public void Save(string json)
            {
                using(var sw=new StreamWriter(json))
                {
                    sw.Write(Json.stringify(this));
                }
            }
        }

        /// <summary>
        /// convert word to png
        /// </summary>
        /// <param name="src">the word file path</param>
        /// <param name="dst_f">the destinace out file path format path with page number ,like xxx-{0}.png</param>
        /// <param name="log">whether show detail log</param>
        /// <returns>the numver of page</returns>
        public static Res word2img(String src, String dst_f, bool log = false)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var pages = 0;
            L.D("executing word2png by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new word.Application();
            try
            {
                app.Visible = true;
                var doc = app.Documents.Open(as_src, false, true);
                doc.ShowGrammaticalErrors = false;
                //doc.ShowRevisions = false;
                doc.ShowSpellingErrors = false;
                if (doc.Windows.Count < 1)
                {
                    L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", as_src, as_dst_f, pages);
                    doc.Close(false);
                    res.Code = 404;
                    return res;
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
                    L.D("executing word2png by file({0}),destination format({1}) with {2} page found", as_src, as_dst_f, pages);
                }
                for (var i = 1; i <= pane.Pages.Count; i++)
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
                    catch (Exception e)
                    {
                        break;
                    }
                    using (var ms = new MemoryStream((byte[])(bits)))
                    {
                        Image.FromStream(ms).Save(spath, ImageFormat.Png);
                    }
                    res.Files.Add(String.Format(dst_f, pages));
                    pages += 1;
                }
                //  panes += 1;
                //}
                //}
                L.D("executing word2png by file({0}),destination format({1}) done with pages({2})", as_src, as_dst_f, pages);
                doc.Close(false);
                res.Code = 0;
                res.Count = pages;
            }
            catch (Exception e)
            {
                L.E(e, "executing word2png by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f, e.Message);
                res.Code = 500;
                throw e;
            }
            finally
            {
                app.Quit();
            }
            return res;
        }

        /// <summary>
        /// convert excel to pdf
        /// </summary>
        /// <param name="src">the excel file path</param>
        /// <param name="dst_f">the destinace out file path format path with sheet number ,like xxx-{0}.pdf</param>
        /// <param name="log">whether show detail log</param>
        /// <returns>the number of sheets</returns>
        public static Res excel2pdf(String src, String dst_f, bool log = false)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var sheets = 0;
            L.D("executing excel2pdf by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new excel.Application();
            try
            {
                app.Visible = true;
                var books = app.Workbooks.Open(as_src, 0, true, 5, "", "",
                    true, excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
                        L.D("excel2pdf parsing file({0},{1}) to {3}", as_src, sheets, spath);
                    }
                    sheet.ExportAsFixedFormat(excel.XlFixedFormatType.xlTypePDF, spath);
                    res.Files.Add(String.Format(dst_f, sheets));
                    sheets += 1;
                }
                books.Close(true, null, null);
                L.D("executing excel2pdf by file({0}),destination format({1}) done with sheets({2})", as_src, as_dst_f, sheets);
                res.Code = 0;
                res.Count = sheets;
            }
            catch (Exception e)
            {
                L.E(e, "executing excel2pdf by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f);
                res.Code = 500;
                throw e;
            }
            finally
            {
                app.Quit();
            }
            return res;
        }

        /// <summary>
        /// convert ppt to pdf
        /// </summary>
        /// <param name="src">the ppt file path</param>
        /// <param name="dst_f">the destinace out file path format path with sheet number ,like xxx-{0}.png</param>
        /// <param name="filterName">the image filter name</param>
        /// <param name="scaleWidth">the image width</param>
        /// <param name="scaleHeight">the image height</param>
        /// <param name="log">whether show detail log</param>
        /// <returns>the number of slides</returns>
        public static Res ppt2img(String src, String dst_f, string filterName = "png", int scaleWidth = 0, int scaleHeight = 0, bool log = false)
        {
            //ILog L = Log.New();
            var res = new Res(src);
            var as_src = Path.GetFullPath(src);
            var as_dst_f = Path.GetFullPath(dst_f);
            var slides = 0;
            L.D("executing word2png by file({0}),destination format({1})", as_src, as_dst_f);
            var app = new ppt.Application();
            try
            {
                var doc = app.Presentations.Open(as_src, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                foreach (ppt.Slide slide in doc.Slides)
                {
                    var spath = String.Format(as_dst_f, slides);
                    if (log)
                    {
                        L.D("word2png parsing file({0},{1}) to {2}", as_src, slides, spath);
                    }
                    slide.Export(spath, filterName, scaleWidth, scaleHeight);
                    res.Files.Add(String.Format(dst_f, slides));
                    slides += 1;

                }
                L.D("executing word2png by file({0}),destination format({1}) done with slides({2})", as_src, as_dst_f, slides);
                res.Code = 0;
                res.Count = slides;
            }
            catch (Exception e)
            {
                L.E("executing word2png by file({0}),destination format({1}) done with error->{2}", as_src, as_dst_f, e.Message);
                res.Code = 500;
                throw e;
            }
            finally
            {
                app.Quit();
            }
            return res;
        }
    }
}
