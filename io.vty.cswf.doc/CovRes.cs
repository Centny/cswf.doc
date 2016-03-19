using io.vty.cswf.util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace io.vty.cswf.doc
{
    /// <summary>
    /// the struct of result
    /// </summary>
    [DataContract]
    public class CovRes
    {
        /// <summary>
        /// the result code, 0 is success, other is fail.
        /// </summary>
        [DataMember(Name = "code")]
        public int Code
        {
            get; set;
        }
        /// <summary>
        /// the result count.
        /// </summary>
        [DataMember(Name = "count")]
        public int Count
        {
            get; set;
        }
        /// <summary>
        /// the result data.
        /// </summary>
        [DataMember(Name = "files")]
        public IList<string> Files
        {
            get; set;
        }
        /// <summary>
        /// the souce file
        /// </summary>
        [DataMember(Name = "src")]
        public string Src
        {
            get; set;
        }

        public CovRes()
        {
            this.Files = new List<string>();
        }
        /// <summary>
        /// constructor by souce file path.
        /// </summary>
        /// <param name="src">the source file path</param>
        public CovRes(string src)
        {
            this.Src = src;
            this.Files = new List<string>();
        }
        /// <summary>
        /// saving the result to file with json format.
        /// </summary>
        /// <param name="json"></param>
        public void Save(string json)
        {
            using (var sw = new StreamWriter(json))
            {
                sw.Write(Json.stringify(this));
            }
        }
        public void Trim(string prefix)
        {
            for (var i = 0; i < this.Files.Count; i++)
            {
                this.Files[i] = this.Files[i].Replace(prefix, "");
            }
        }
    }
}
