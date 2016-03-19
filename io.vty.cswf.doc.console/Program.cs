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
        static readonly ILog L = Log.New(); //for inintial logger.
        static void Usage()
        {
            Console.WriteLine("Usage:\n" +
                "   cswf-doc <options> -w <word file> <image path format>   convert word document to image\n" +
                "   cswf-doc <options> -e <excel file> <pdf path format>    convert excel doucmnet to pdf\n" +
                "   cswf-doc <options> -p <ppt file> <image path format>    convert ppt doucmnet to image\n" +
                "   cswf-doc <options> -pdf <pdf file> <image path format>    convert pdf doucmnet to image\n" +
                "   cswf-doc <options> -i <image> <image path format>    convert image to image\n" +
                "   \n" +
                " common options:\n" +
                "   -l show detail log\n" +
                "   -o <json result>\n" +
                "   -prefix <the trim prefix to out json>" +
                "   -w <width>  the output image width\n" +
                "   -h <height>  the output image height\n" +
                "   -dx <density x>  the read pdf density x\n" +
                "   -dy <density y>  the read pdf density y\n" +
                "   -beg <begin>  the begin int\n" +
                "   -maxc <maximum concurrency>  the maximum concurrency to convert\n"
                );
        }
        static void Main(string[] args)
        {
            try
            {
                var cargs = Args.parseArgs(new string[] { "w", "e", "p", "l", "pdf" }, args, 0);
                var log = cargs.Exist("l");
                var w = cargs.Exist("w");
                var e = cargs.Exist("e");
                var p = cargs.Exist("p");
                var pdf = cargs.Exist("pdf");
                var i = cargs.Exist("i");
                int width, height;
                cargs.IntVal("W", out width, 768);
                cargs.IntVal("h", out height, 1024);
                double densityx, densityy;
                cargs.DoubleVal("dx", out densityx, 96);
                cargs.DoubleVal("dy", out densityy, 96);
                int beg = 0;
                cargs.IntVal("beg", out beg, 0);
                int maxc;
                cargs.IntVal("maxc", out maxc, 16);
                TaskPool.Shared.MaximumConcurrency = maxc;
                if (cargs.Vals.Count < 2)
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
                L.D("cswf-doc is starting...");
                CovProc cov;
                if (w)
                {
                    cov = new WordCov(cargs.Vals[0], cargs.Vals[1], width, height, beg);
                }
                else if (e)
                {
                    cov = new ExcelCov(cargs.Vals[0], cargs.Vals[1], width, height, densityx, densityy, beg);
                }
                else if (p)
                {
                    cov = new PowerPointCov(cargs.Vals[0], cargs.Vals[1], beg);
                }
                else if (pdf)
                {
                    cov = new PdfCov(cargs.Vals[0], cargs.Vals[1], width, height, densityx, densityy, beg);
                }
                else if (i)
                {
                    cov = new ImgCov(cargs.Vals[0], cargs.Vals[1], width, height, beg);
                }
                else
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
                cov.ShowLog = log;
                cov.Exec();
                cov.Dispose();
                var res = cov.Result;
                string prefix = "";
                if (cargs.StringVal("prefix", out prefix))
                {
                    res.Trim(prefix);
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
                L.E(e, "do converter error by args({0})->", String.Join(",", args), e.Message);
                Environment.ExitCode = 1;
            }
            finally
            {
                WordCov.Cached.Clear();
                ExcelCov.Cached.Clear();
            }
            L.D("do converter done...");
        }
    }
}
