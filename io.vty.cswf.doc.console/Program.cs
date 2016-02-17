using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc.console
{
    class Program
    {
        static void Main(string[] args)
        {
            //new Converter().excel2pdf("C:\\xxx\\xx.xlsx", "C:\\xxx\\xx-{0}.pdf", true);
            new Converter().ppt2img("C:\\xxx\\xx.pptx", "C:\\xxx\\ppt-{0}.png");
        }
    }
}
