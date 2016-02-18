using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using io.vty.cswf.util;
using System.IO;
using io.vty.cswf.log;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace io.vty.cswf.doc.console
{
    class Program
    {
        static void Usage()
        {
            Console.WriteLine("Usage:\n" +
                "   cswf-doc -l -o <json result> -w <word file> <image path format>   convert word document to image\n" +
                "   cswf-doc -l -o <json result> -e <excel file> <pdf path format>   convert excel doucmnet to pdf\n" +
                "   cswf-doc -l -o <json result> -f <filter name> -W <image width> -H <image height> -p <ppt file> <image path format>    convert ppt doucmnet to image\n"
                );
        }
        static void Main(string[] args)
        {
            ILog L = Log.New();//for inintial logger.
            try
            {
                var cargs = Args.parseArgs(new string[] { "w", "e", "p", "l" }, args, 0);
                var log = cargs.Exist("l");
                var w = cargs.Exist("w");
                var e = cargs.Exist("e");
                var p = cargs.Exist("p");
                string filter = "";
                int width, height;
                cargs.StringVal("f", out filter);
                cargs.IntVal("W", out width);
                cargs.IntVal("h", out height);
                if (cargs.Vals.Count < 2)
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
                Converter.Res res;
                if (w)
                {
                    res = Converter.word2img(cargs.Vals[0], cargs.Vals[1], log);
                }
                else if (e)
                {
                    res = Converter.excel2pdf(cargs.Vals[0], cargs.Vals[1], log);
                }
                else if (p)
                {
                    res = Converter.ppt2img(cargs.Vals[0], cargs.Vals[1], filter, width, height, log);
                }
                else
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
                string json = "";
                if (cargs.StringVal("o", out json))
                {
                    L.D("converter saving result to json file({0})", json);
                    res.Save(json);
                }
                Environment.ExitCode = 0;
            }
            catch (Exception e)
            {
                L.E(e, "do converter error by args({0})", String.Join(",", args));
                Environment.ExitCode = 1;
            }
            L.D("do converter done...");
        }
    }
}
