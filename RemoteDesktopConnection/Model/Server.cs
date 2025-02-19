﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RemoteDesktopConnection.Model
{
    public class Server
    {
        public string Name { get; set; }
        public bool IsMaintenace { get; set; }
        public int Index { get; set; }
        public string Administrator { get; set; }

        public Server(string name, bool isMaintenace, string administrator)
        {
            Name = name;
            IsMaintenace = isMaintenace;
            Administrator = administrator;
            Index = Convert.ToInt32(Regex.Split(name, "AGMS")[1]);
        }
    }
}
