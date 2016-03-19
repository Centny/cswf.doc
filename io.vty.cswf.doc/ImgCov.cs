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
        }

        public override void Exec()
        {
            try
            {
                var img = new ImageMagick.MagickImage(this.Src);
                img.BackgroundColor = new MagickColor(Color.White);
                if (".jpg".Equals(Path.GetExtension(this.DstF), StringComparison.OrdinalIgnoreCase))
                {
                    img.HasAlpha = false;
                }
                img.Resize(this.size);
                img.Write(String.Format(this.AsDstF, this.Beg));
                this.Result.Files.Add(String.Format(this.DstF, this.Beg));
            }
            catch (Exception e)
            {
                this.Result.Code = 500;
                this.Fails.Add(e);
            }
        }
    }
}
