using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RemoteDesktopConnection.Model
{
    public class Software
    {
        public string Server { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }

        public Software(
            string server,
            string name,
            string version)
        {
            Server = server;
            Name = name;
            Version = version;
        }
    }
}
