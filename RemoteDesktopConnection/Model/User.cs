using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RemoteDesktopConnection.Model
{
    public class User
    {
        public string Server { get; set; }
        public string Name { get; set; }
        public string Date { get; set; }
        public string TaskDate { get; set; }
        public DateTime DateOriginal { get; set; }
        public DateTime TaskDateriginal { get; set; }
        public bool IsLogged { get; set; }
        public bool HasTaskFinished { get; set; }
        public string Email { get; set; }

        public User(
            string server,
            string name,
            string date,
            string taskDate,
            DateTime dateOriginal,
            DateTime taskDateOriginal,
            bool isLogged,
            bool hasTaskFinished,
            string email)
        {
            Server = server;
            Name = name;
            Date = date;
            TaskDate = taskDate;
            DateOriginal = dateOriginal;
            TaskDateriginal = taskDateOriginal;
            IsLogged = isLogged;
            HasTaskFinished = hasTaskFinished;
            Email = email;
        }
    }
}
