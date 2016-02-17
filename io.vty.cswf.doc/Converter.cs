using io.vty.cswf.log;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using ppt = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace io.vty.cswf.doc
{
    public class Converter
    {
        private static readonly ILog L = Log.New();

        /// <summary>
        /// convert word to png
        /// </summary>
        /// <param name="src">the word file path</param>
        /// <param name="dst_f">the destinace out file path format path with page number ,like xxx-{0}.png</param>
        /// <param name="format">the image format</param>
        /// <param name="log">whether show detail log</param>
        /// <returns>the numver of page</returns>
        public int word2img(String src, String dst_f,ImageFormat format, bool log=false)
        {
            var panes = 0;
            var pages = 0;
            L.D("executing word2png by file({0}),destination format({1})", src, dst_f);
            var app = new word.Application();
            app.Visible = true;
            var doc = app.Documents.Open(src, false, true);
            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;
            foreach (word.Window window in doc.Windows)
            {
                foreach (word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var spath = String.Format(dst_f, pages);
                        if (log)
                        {
                            L.D("word2png parsing file({0},{1},{2}) to {3}", src, panes, pages, spath);
                        }
                        var page = pane.Pages[i];
                        var bits = page.EnhMetaFileBits;
                        using (var ms = new MemoryStream((byte[])(bits)))
                        {
                            Image.FromStream(ms).Save(spath, format);
                        }
                        pages += 1;
                    }
                    panes += 1;
                }
            }
            doc.Close(false);
            app.Quit();
            L.D("executing word2png by file({0}),destination format({1}) done with panes({2}),pages({3})", src, dst_f, panes, pages);
            return pages;
        }
        /// <summary>
        /// convert excel to pdf
        /// </summary>
        /// <param name="src">the excel file path</param>
        /// <param name="dst_f">the destinace out file path format path with page number ,like xxx-{0}.pdf</param>
        /// <param name="log">whether show detail log</param>
        /// <param name="sheets">the number of sheets</param>
        public int excel2pdf(String src, String dst_f, bool log=false)
        {
            var sheets = 0;
            L.D("executing excel2pdf by file({0}),destination format({1})", src, dst_f);
            var app = new excel.Application();
            app.Visible = true;
            var books = app.Workbooks.Open(src, 0, true, 5, "", "",
                true, excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (excel.Worksheet sheet in books.Worksheets)
            {
                var range = sheet.UsedRange;
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                if (rows < 2 && cols < 2)
                {
                    Object text = range.Text;
                    if(text is string&& ((string)text).Length < 1)
                    {
                        continue;
                    }
                }
                var spath = String.Format(dst_f, sheets);
                if (log)
                {
                    L.D("excel2pdf parsing file({0},{1}) to {3}", src, sheets, spath);
                }
                sheet.ExportAsFixedFormat(excel.XlFixedFormatType.xlTypePDF, spath);
                sheets += 1;
            }
            books.Close(true, null, null);
            app.Quit();
            L.D("executing excel2pdf by file({0}),destination format({1}) done with sheets({2})", src, dst_f, sheets);
            return sheets;
        }
        public int ppt2img(String src, String dst_f, string filterName="png", int scaleWidth = 0, int scaleHeight = 0, bool log=false)
        {
            var slides = 0;
            L.D("executing word2png by file({0}),destination format({1})", src, dst_f);
            var app = new ppt.Application();
            var doc = app.Presentations.Open(src, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            foreach (ppt.Slide slide in doc.Slides)
            {
                var spath = String.Format(dst_f, slides);
                if (log)
                {
                    L.D("word2png parsing file({0},{1}) to {2}", src, slides, spath);
                }
                slide.Export(spath, filterName, scaleWidth, scaleHeight);

                slides += 1;

            }
            app.Quit();
            L.D("executing word2png by file({0}),destination format({1}) done with slides({2})", src, dst_f, slides);
            return slides;
        }
    }
}
