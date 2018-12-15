using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Cet4_6Tools.Models
{
    public class Clazz
    {
        public string folder { get; set; }
        public int count { get; set; }

        public Clazz(string folder, int count)
        {
            this.folder = folder;
            this.count = count;
        }

        public Clazz()
        {
        }
    }
}