using io.vty.cswf.log;
using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

//[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace io.vty.cswf.doc.mtest
{
    class Program
    {
        //static readonly ILog L = Log.New(); //for inintial logger.
        static void Main(string[] args)
        {
            //L.D("start run...");
            var running = false;
            new Thread(() =>
            {
                while (true)
                {
                    if (!running)
                    {
                        Thread.Sleep(3000);
                        continue;
                    }
                    WordCov cov = new WordCov("test\\xx.docx", "docx-{0}.jpg");
                    cov.Exec();
                }
                Console.WriteLine("exit...");
            }).Start();
            WordCov.Cached.MaxIdle = 0;
            ProcKiller.AddName("WINWORD");
            ProcKiller.StartTimer(5000);
            while (true)
            {
                var line=Console.ReadLine();
                running = "r".Equals(line);
                if ("e".Equals(line))
                {
                    break;
                }
            }
        }
    }
}
