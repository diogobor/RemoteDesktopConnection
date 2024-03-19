using Google.Apis.Gmail.v1;
using Google.Apis.Sheets.v4;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.Win32;
using RemoteDesktopConnection.Control.Database;
using RemoteDesktopConnection.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RemoteDesktopConnection.Control
{
    public static class Management
    {
        /// <summary>
        /// public variables
        /// </summary>
        public static List<Model.User> Users_agms2 = new();
        public static List<Model.User> Users_agms3 = new();
        public static List<Model.User> Users_agms4 = new();
        public static List<Model.Software> Software_agms2 = new();
        public static List<Model.Software> Software_agms3 = new();
        public static List<Model.Software> Software_agms4 = new();
        public const int TIMEOUT_HOURS_USER_CONNECTION = 0;
        public const int INTERVAL_TIME_REFRESH_SECONDS = 100;
        public static DateTime last_refresh;
        public static Process ConnectServerProcess;
        public static bool IsConnected { get; set; } = false;
        public static DateTime Connected_initial_time;
        public static int CurrentPID { get; set; } = -1;
        public static string IP_address { get; set; } = string.Empty;

        /// <summary>
        /// private variables
        /// </summary>
        private static Model.Database Database { get; set; }
        private static CultureInfo cultureInfo = new CultureInfo("en-US");
        private const int TIMEOUT_DAYS_USER_TASKS = 10;

        public static Model.Database GetDatabase()
        {
            return Database;
        }

        public static void SetDatabase()
        {
            string googleClientID = "680087468736-skuel9mfejmals227c7kqrge7nm89qnf.apps.googleusercontent.com";
            string googleClientSecret = "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";
            string spreadsheetId = "1ffT7OH5LbFQjaAkjx54CuzQfjIUX6j_rH0WUXzD9tKY";
            Database = new Model.Database(googleClientID, googleClientSecret, spreadsheetId);
        }

        public static string GetUser()
        {
            return System.Environment.UserName.ToLower().ToString();
        }

        public static List<(string user, string email, string server)> CheckLongerUsers()
        {
            List<(string user, string email, string server)> user_to_send_email = new();
            DateTime currentDate = DateTime.Now;

            #region AGMS2
            List<Model.User> userslogged = Users_agms2.Where(a => a.IsLogged).ToList();

            foreach (var dgi in userslogged)
            {
                DateTime user_date = DateTime.ParseExact(dgi.Date, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.Days > 0 || difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                    user_to_send_email.Add((dgi.Name, dgi.Email, "AGMS2"));
            }
            #endregion

            #region AGMS3
            userslogged = Users_agms3.Where(a => a.IsLogged).ToList();

            foreach (var dgi in userslogged)
            {
                DateTime user_date = DateTime.ParseExact(dgi.Date, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.Days > 0 || difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                    user_to_send_email.Add((dgi.Name, dgi.Email, "AGMS3"));
            }
            #endregion

            #region AGMS4
            userslogged = Users_agms4.Where(a => a.IsLogged).ToList();

            foreach (var dgi in userslogged)
            {
                DateTime user_date = DateTime.ParseExact(dgi.Date, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.Days > 0 || difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                    user_to_send_email.Add((dgi.Name, dgi.Email, "AGMS4"));
            }
            #endregion

            return user_to_send_email.Distinct().ToList();
        }

        public static List<(string user, string email, string server)> CheckLongerTasks()
        {
            List<(string user, string email, string server)> user_to_send_email = new();
            DateTime currentDate = DateTime.Now;

            #region AGMS2
            List<Model.User> tasksNotFinished = Users_agms2.Where(a => !a.HasTaskFinished).ToList();

            foreach (var user in tasksNotFinished)
            {
                DateTime user_date = DateTime.ParseExact(user.TaskDate, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.TotalDays > TIMEOUT_DAYS_USER_TASKS)
                    user_to_send_email.Add((user.Name, user.Email, "AGMS2"));
            }
            #endregion

            #region AGMS3
            tasksNotFinished = Users_agms3.Where(a => !a.HasTaskFinished).ToList();

            foreach (var dgi in tasksNotFinished)
            {
                DateTime user_date = DateTime.ParseExact(dgi.TaskDate, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.TotalDays > TIMEOUT_DAYS_USER_TASKS)
                    user_to_send_email.Add((dgi.Name, dgi.Email, "AGMS3"));
            }
            #endregion

            return user_to_send_email.Distinct().ToList();
        }

        public static void ConnectAGMS(string username, string password, out string error)
        {
            try
            {
                if (String.IsNullOrEmpty(IP_address))
                    throw new Exception("ERROR: Invalid ip address!");
                NotAskForPwdOnRDC();

                var process = new Process();
                process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdKey.exe");
                process.StartInfo.Arguments = String.Format(@"/generic:{0} /user:{1} /pass:{2}", IP_address, username, password);
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();

                ConnectServerProcess = new Process();
                ConnectServerProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                ConnectServerProcess.StartInfo.Arguments = String.Format(@"/f /v:{0}", IP_address);
                ConnectServerProcess.EnableRaisingEvents = true;
                ConnectServerProcess.StartInfo.UseShellExecute = false;
                ConnectServerProcess.StartInfo.RedirectStandardOutput = true;
                ConnectServerProcess.Start();
                ConnectServerProcess.PriorityClass = ProcessPriorityClass.High;
                ConnectServerProcess.PriorityBoostEnabled = true;
                IsConnected = true;

                Connected_initial_time = ConnectServerProcess.StartTime;
                CurrentPID = ConnectServerProcess.Id;

                var outputResultPromise = ConnectServerProcess.StandardOutput.ReadToEndAsync();
                outputResultPromise.ContinueWith(o =>
                IsConnected = false
                );
                error = "";

            }
            catch (Exception e)
            {
                error = "";
            }
        }

        public static bool CheckProcessStatus()
        {
            var GetAllTasks = new Process();
            GetAllTasks.StartInfo.FileName = "tasklist";
            GetAllTasks.StartInfo.Arguments = "/fo CSV";
            GetAllTasks.EnableRaisingEvents = true;
            GetAllTasks.StartInfo.UseShellExecute = false;
            GetAllTasks.StartInfo.RedirectStandardOutput = true;
            GetAllTasks.StartInfo.CreateNoWindow = true;
            GetAllTasks.Start();
            GetAllTasks.PriorityClass = ProcessPriorityClass.High;
            GetAllTasks.PriorityBoostEnabled = true;

            var all_tasks = GetAllTasks.StandardOutput.ReadToEnd();
            all_tasks = all_tasks.Replace("\"", "");
            GetAllTasks.WaitForExit();

            List<int> all_PIDs = new();
            string[] cols = Regex.Split(all_tasks, "\r\n");
            foreach (string col in cols)
            {
                if (String.IsNullOrEmpty(col) || col.StartsWith("Image Name")) continue;

                string[] tasks = Regex.Split(col, ",");
                all_PIDs.Add(int.Parse(tasks[1]));
            }
            all_PIDs.Sort();

            if (all_PIDs.Contains(CurrentPID))
                return true;
            return false;
        }


        private static void NotAskForPwdOnRDC()
        {
            try
            {
                // Specify the path to the key containing the settings for Remote Desktop Connection.
                string keyPath = @"Software\Microsoft\Terminal Server Client";
                // Create or open the registry key.
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true))
                {
                    if (key != null)
                    {
                        // Set the value of the "PromptForCredentials" DWORD to 0 to disable the "Always ask for credentials" option.
                        key.SetValue("PromptForCredentials", 0, RegistryValueKind.DWord);
                        Console.WriteLine("Successfully disabled 'Always ask for credentials' option.");
                    }
                    else
                    {
                        Console.WriteLine("Registry key not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
