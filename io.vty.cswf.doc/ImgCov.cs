using ImageMagick;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class ImgCov : CovProc
    {
        protected MagickGeometry size;
        public ImgCov(String src, String dst_f, int maxw = 768, int maxh = 1024, int beg = 0) : base(src, dst_f, maxw, maxh, beg)
        {
            this.size = new MagickGeometry(maxw, maxh);
            this.size.Greater = true;
        }

        public override void Exec()
        {
            try
            {
                var img = new ImageMagick.MagickImage(this.AsSrc);
                img.BackgroundColor = new MagickColor(Color.White);
                if (".jpg".Equals(Path.GetExtension(this.DstF), StringComparison.OrdinalIgnoreCase))
                {
                    img.HasAlpha = false;
                }
                img.Resize(this.size);
                var as_dst = String.Format(this.AsDstF, this.Beg);
                var as_dir = Path.GetDirectoryName(as_dst);
                if (!Directory.Exists(as_dir))
                {
                    Directory.CreateDirectory(as_dir);
                }
                img.Write(as_dst);
                this.Result.Files.Add(String.Format(this.DstF, this.Beg));
                this.Result.Count += 1;
            }
            catch (Exception e)
            {
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
        }
    }
}
