using ImageMagick;
using io.vty.cswf.log;
using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class PdfCov : CovProc
    {
        private static readonly ILog L = Log.New();
        public double DensityX { get; protected set; }
        public double DensityY { get; protected set; }
        protected MagickGeometry size;
        protected MagickReadSettings settings;
        public PdfCov(String src, String dst_f, int maxw = 768, int maxh = 1024, double densityx = 96, double densityy = 96, int beg = 0) : base(src, dst_f, maxw, maxh, beg)
        {

            this.DensityX = densityx;
            this.DensityY = densityy;
            this.size = new MagickGeometry(maxw, maxh);
            this.size.Greater = true;
            this.settings = new MagickReadSettings();
            this.settings.Density = new PointD(this.DensityX, this.DensityY);
        }
        public override void Exec()
        {
            //var pages = this.Beg;
            L.D("executing pdf2img by file({0}),destination format({1})", this.AsSrc, this.AsDstF);
            this.Cdl.add();
            MagickImageCollection images = new MagickImageCollection();
            try
            {
                this.Pdf2imgProc(images, this.AsSrc, -1, 0);
            }
            catch (Exception e)
            {
                L.E(e, "executing pdf2img by file({0}),destination format({1}) fail with error->{2}", this.AsSrc, this.AsDstF, e.Message);
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
            this.Cdl.done();
            this.Cdl.wait();
            images.Dispose();
            L.D("executing pdf2img by file({0}),destination format({1}) done with pages({2}),fails({3})", this.AsSrc, this.AsDstF, this.Result.Count, this.Fails.Count);
        }

        protected virtual int Pdf2imgProc(MagickImageCollection images, String pdf, int idx, int file_c)
        {
            images.Read(pdf, settings);
            int pages = images.Count;
            if (idx < 0)
            {
                this.Total = new int[pages];
                this.Done = new int[pages];
                Util.set(this.Total, 1);
                Util.set(this.Done, 0);
            }
            else
            {
                this.Total[idx] = pages;
            }
            for (var i = 0; i < pages; i++)
            {
                this.Pdf2imgProc(images[i], idx, i, file_c);
            }
            return pages;
        }
        protected virtual void Pdf2imgProc(MagickImage image, int idx, int i, int file_c)
        {
            this.Cdl.add();
            this.Result.Files.Add("");
            TaskPool.Queue(args =>
            {
                this.RunPdf2imgProc(image, idx, i, file_c);
            }, 0);
        }
        private void RunPdf2imgProc(MagickImage image, int idx, int i, int file_c)
        {
            try
            {
                var as_dst = String.Format(this.AsDstF, this.Beg + file_c + i);
                var as_dir = Path.GetDirectoryName(as_dst);
                if (!Directory.Exists(as_dir))
                {
                    Directory.CreateDirectory(as_dir);
                }
                if (this.ShowLog)
                {
                    L.D("pdf2img parsing file({0},{1}) to {2}", this.AsSrc, file_c + i, as_dst);
                }
                image.BackgroundColor = new MagickColor(Color.White);
                image.HasAlpha = false;
                image.Resize(size);
                image.Write(as_dst);
                this.Result.Count++;
                this.Result.Files[file_c + i] = String.Format(this.DstF, this.Beg + file_c + i);
                if (idx < 0)
                {
                    this.Done[i] = 1;
                }
                else
                {
                    this.Done[idx] += 1;
                }
                this.OnDone();
            }
            catch (Exception e)
            {
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
            this.Cdl.done();
        }
    }
}
