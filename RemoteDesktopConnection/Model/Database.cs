using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RemoteDesktopConnection.Model
{
    public class Database
    {
        public string GoogleClientID { get; set; }
        public string GoogleClientSecret { get; set; }
        public string SpreadsheetID { get; set; }

        public Database(string googleClientID, string googleClientSecret, string spreadsheetID)
        {
            GoogleClientID = googleClientID;
            GoogleClientSecret = googleClientSecret;
            SpreadsheetID = spreadsheetID;
        }
    }
}
