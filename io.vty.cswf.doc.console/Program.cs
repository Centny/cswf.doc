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
                "   cswf-doc <options> -w <word file> <image path format>   convert word document to image\n" +
                "   cswf-doc <options> -e <excel file> <pdf path format>    convert excel doucmnet to pdf\n" +
                "   cswf-doc <options> -p <ppt file> <image path format>    convert ppt doucmnet to image\n" +
                "   cswf-doc <options> -x <source file> <destiance format>  exec something\n" +
                "   \n" +
                " common options:\n" +
                "   -l show detail log\n" +
                "   -o <json result>\n" +
                "   -prefix <the trim prefix to out json>" +
                "   \n" +
                " ppt options:\n" +
                "   -f <filer name> the filter nam to output file, like png\n" +
                "   -W <width>  the output image width\n" +
                "   -H <height>  the output image height\n"
                );
        }
        static void Main(string[] args)
        {
            TaskPool.Shared.MaximumConcurrency = 10;
            WordCov cov = new WordCov("test\\xx.docx", "docx-{0}.jpg");
            cov.Exec();
            foreach (var e in cov.Fails)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.WriteLine("--->\n");
            }
            //Assert.AreEqual(0, cov.Fails.Count);
            WordCov.Cached.Clear();
            /*
            ILog L = Log.New();//for inintial logger.
            try
            {
                var cargs = Args.parseArgs(new string[] { "w", "e", "p", "l", "x" }, args, 0);
                var log = cargs.Exist("l");
                var w = cargs.Exist("w");
                var e = cargs.Exist("e");
                var p = cargs.Exist("p");
                var x = cargs.Exist("x");
                string filter = "";
                int width, height;
                cargs.StringVal("f", out filter, "png");
                cargs.IntVal("W", out width);
                cargs.IntVal("h", out height);
                string exe_c = "", exe_f = "", exe_a = "", exe_p = "";
                cargs.StringVal("exe_c", out exe_c);
                cargs.StringVal("exe_f", out exe_f);
                cargs.StringVal("exe_a", out exe_a);
                cargs.StringVal("exe_p", out exe_p);
                int beg = 0;
                if (cargs.Vals.Count < 2)
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
                L.D("using exe by \n exe_c:{0}\n exe_f:{1}\n exe_a:{2}", exe_c, exe_f, exe_a);
                Converter.OnProcess proc = null;
                if (exe_c.Length > 0 && exe_f.Length > 0)
                {
                    var exec = new Converter.Proc(exe_c, exe_f);
                    exec.Progress = exe_p;
                    exec.Args = exe_a;
                    proc = exec.exec;
                }
                Converter.Res res;
                if (w)
                {
                    res = Converter.word2img(cargs.Vals[0], cargs.Vals[1], beg, log, proc);
                }
                else if (e)
                {
                    res = Converter.excel2pdf(cargs.Vals[0], cargs.Vals[1], beg, log, proc);
                }
                else if (p)
                {
                    res = Converter.ppt2img(cargs.Vals[0], cargs.Vals[1], beg, filter, width, height, log, proc);
                }
                else if (x)
                {
                    res = Converter.exec(cargs.Vals[0], cargs.Vals[1], beg, log, proc);
                }
                else
                {
                    Usage();
                    Environment.ExitCode = 1;
                    return;
                }
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
                L.E(e, "do converter error by args({0})", String.Join(",", args));
                Environment.ExitCode = 1;
            }
            L.D("do converter done...");
            */
        }
    }
}
