using io.vty.cswf.log;
using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    public class Samba
    {
        public delegate void OnFail(Samba s, Exception e);
        public delegate void OnSuccess(Samba s);
        public static readonly ILog L = Log.New();
        public const int READ = 1;
        public const int RW = 2;
        public string Volume { get; set; }
        public string Uri { get; set; }
        public string User { get; set; }
        public string Pwd { get; set; }
        public IDictionary<string, int> Paths { get; set; }
        public OnFail Fail { get; set; }
        public OnSuccess Success { get; set; }
        public bool Activated { get; set; }

        public Samba(string volume, string uri)
        {
            this.Volume = volume;
            this.Uri = uri;
            this.Paths = new Dictionary<string, int>();
        }
        public int Remount(out string res)
        {
            L.I("Samba remount by volume({0})/uri({1})/user({2})/pwd({3})", this.Volume, this.Uri, this.User, this.Pwd);
            Exec.exec(out res, "net", "use", this.Volume, "/delete", "/y");
            if (String.IsNullOrEmpty(this.User))
            {
                return Exec.exec(out res, "net", "use", this.Volume, this.Uri);
            }
            else
            {
                return Exec.exec(out res, "net", "use", "/user:" + this.User, this.Volume, this.Uri, this.Pwd);
            }
        }

        public void Test()
        {
            foreach (var path in this.Paths)
            {
                switch (path.Value)
                {
                    case READ:
                        //L.D("Samba test path({0}) by READ", path.Value);
                        Util.read(path.Key);
                        break;
                    case RW:
                        //L.D("Samba test path({0}) by RW", path.Value);
                        Util.write(path.Key, "key");
                        Util.read(path.Key);
                        break;
                }
            }
        }

        public int Check(bool retry = false)
        {
            try
            {
                this.Test();
                //L.D("Samba test path success");
                if (this.Success != null)
                {
                    this.Success(this);
                }
                this.Activated = true;
                return 0;
            }
            catch (Exception e)
            {
                L.E("Samba test path({0}) fail with error({1}),will try remount", this.Paths, e.Message);
                if (this.Activated && this.Fail != null)
                {
                    this.Fail(this, e);
                }
                this.Activated = false;
                if (retry)
                {
                    return -1;
                }
            }
            string res;
            var code = this.Remount(out res);
            if (code != 0)
            {
                L.E("Samba try remount fail with code({0}),result(\n{1}\n)", code, res);
                return code;
            }
            if (!retry)
            {
                return this.Check(true);
            }
            else
            {
                return code;
            }
        }

        public static readonly IList<Samba> Volumes = new List<Samba>();
        public static int RunDelay = 1000;
        public static int ChkDelay = 6000;
        public static int Next = 0;
        public static bool Running = false;

        public static void LoopChecker()
        {
            Running = true;
            int delay = 0;
            while (Running)
            {
                if (delay >= ChkDelay || Next > 0)
                {
                    foreach (var samba in Volumes)
                    {
                        samba.Check();
                    }
                    delay = 0;
                    Next = 0;
                }
                Thread.Sleep(RunDelay);
                delay += RunDelay;
            }
        }

        public static Samba AddVolume(string volume, string uri, string user = null, string pwd = null, IDictionary<string, int> paths = null)
        {
            var samba = new Samba(volume, uri);
            samba.User = user;
            samba.Pwd = pwd;
            samba.Paths = paths;
            Volumes.Add(samba);
            return samba;
        }
        public static Samba AddVolume2(string volume, string uri, string user = null, string pwd = null, string paths = null)
        {
            var samba = new Samba(volume, uri);
            samba.User = user;
            samba.Pwd = pwd;
            if (!String.IsNullOrEmpty(paths))
            {
                samba.Paths = Json.parse<Dictionary<string, int>>(paths);
            }
            Volumes.Add(samba);
            return samba;
        }
    }
}
