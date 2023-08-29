using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Logging;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Runtime;
using System.Security.Cryptography.X509Certificates;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace RemoteDesktopConnection
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public class DataGridItemUser
        {
            public string Server { get; set; }
            public string Name { get; set; }
            public string Date { get; set; }
            public string TaskDate { get; set; }
            public DateTime DateOriginal { get; set; }
            public DateTime TaskDateriginal { get; set; }
            public bool IsLogged { get; set; }
            public bool HasTaskFinished { get; set; }

            public DataGridItemUser(string server, string name, string date, DateTime dateOriginal, string taskDate, DateTime taskDateOriginal, bool isLogged, bool hasTaskFinished)
            {
                Server = server;
                Name = name;
                Date = date;
                TaskDate = taskDate;
                DateOriginal = dateOriginal;
                TaskDateriginal = taskDateOriginal;
                IsLogged = isLogged;
                HasTaskFinished = hasTaskFinished;
            }

        }

        public class DataGridItemSoftware
        {
            public string Server { get; set; }
            public string Name { get; set; }
            public string Version { get; set; }

            public DataGridItemSoftware(string server, string name, string version)
            {
                Server = server;
                Name = name;
                Version = version;
            }
        }

        private System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        private readonly string ApplicationName = "Google Sheet API .NET Quickstart";
        private readonly string SpreadsheetId = "1ffT7OH5LbFQjaAkjx54CuzQfjIUX6j_rH0WUXzD9tKY";
        private SheetsService service_sheets;
        private GmailService service_email;
        private CultureInfo cultureInfo = new CultureInfo("en-US");

        private List<DataGridItemUser> data_agms2_users;
        private List<DataGridItemSoftware> data_agms2_software;
        private List<DataGridItemUser> data_agms3_users;
        private List<DataGridItemSoftware> data_agms3_software;

        private string error_taken_time_connected = "";

        private DateTime last_refresh;

        private System.Windows.Threading.DispatcherTimer dispatcherTimer_checkProcessAlive = new System.Windows.Threading.DispatcherTimer();
        private Process connectServerProcess;
        private bool isConnected { get; set; } = false;
        private DateTime connected_initial_time;
        private int selected_server = -1;

        private const int INTERVAL_TIME_REFRESH_SECONDS = 100;
        private const int OFFSET_TIME_REFRESH_MILLISECONDS = 1000;
        private const int TIMEOUT_HOURS_USER_CONNECTION = 3;
        private const int TIMEOUT_DAYS_USER_TASKS = 10;


        public MainWindow()
        {
            InitializeComponent();
            TBUser.Text = GetUser();
            Init();

            ReadSheets();
            if (data_agms2_users.Count > 0 && data_agms3_users.Count > 0)
            {
                last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }

            this.Closed += new EventHandler(Window_Closed);

            LoadDatagrid();
            SendEmails(CheckLongerTasks(), true);
            SendEmails(CheckLongerUsers(), false);


            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, INTERVAL_TIME_REFRESH_SECONDS);
            dispatcherTimer.Start();

            dispatcherTimer_checkProcessAlive.Tick += new EventHandler(dispatcherTimer_checkProcessAlive_Tick);
            dispatcherTimer_checkProcessAlive.Interval = new TimeSpan(0, 0, 1);

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (connectServerProcess != null && !connectServerProcess.HasExited)
            {
                dispatcherTimer_checkProcessAlive.Stop();
                error_taken_time_connected = "closed";

                connectServerProcess.Kill();
                if (selected_server == -1) return;

                var response = System.Windows.Forms.MessageBox.Show($"Has your task been finished?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (response == System.Windows.Forms.DialogResult.Yes)
                {
                    //Update google sheets
                    UpdateSheet(selected_server, false, true);
                }
                else
                {
                    //Update google sheets
                    UpdateSheet(selected_server, false, false);
                }

                System.Windows.MessageBox.Show(
                            "The server has been closed successfully!",
                            "The Liu Lab :: Information",
                            (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                            (System.Windows.MessageBoxImage)MessageBoxIcon.Information);
            }
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            do
            {
                ReadSheets();
                if (data_agms2_users.Count > 0 && data_agms3_users.Count > 0)
                {
                    last_refresh = DateTime.Now;
                    LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
                }

                if (data_agms2_users.Count == 0 || data_agms3_users.Count == 0)
                {
                    ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));
                    System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                }
            }
            while (data_agms2_users.Count == 0 || data_agms3_users.Count == 0);
            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));

            LoadDatagrid();
        }

        private void dispatcherTimer_checkProcessAlive_Tick(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                Process_Exited(null, null);
                dispatcherTimer_checkProcessAlive.Stop();
                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
            }
            else
            {
                //TimeSpan difference = DateTime.Now - refresh_connected_initial_time;

                //if (difference.Minutes > 5)
                //{
                //    var response = System.Windows.Forms.MessageBox.Show($"Are you still connected?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                //    if (response == System.Windows.Forms.DialogResult.No)
                //    {
                //        if (connectServerProcess != null && !connectServerProcess.HasExited)
                //        {
                //            connectServerProcess.Kill();
                //            Process_Exited(null, null);
                //            dispatcherTimer_checkProcessAlive.Stop();
                //            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
                //        }
                //    }
                //    else if (response == System.Windows.Forms.DialogResult.Yes)
                //        refresh_connected_initial_time = DateTime.Now;
                //}

                TimeSpan difference = DateTime.Now - connected_initial_time;

                if (difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                {
                    connectServerProcess.Kill();

                    //Send an email to the user
                    TBUser.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
                    {
                        SendEmail(TBUser.Text, ("AGMS" + selected_server), false);
                        dispatcherTimer.Start();
                    }));
                    dispatcherTimer_checkProcessAlive.Stop();
                }
            }
        }

        private string GetUser()
        {
            return System.Environment.UserName.ToLower().ToString();
        }

        private static UserCredential Login(string googleClientId, string googleClientSecret, string[] scopes)
        {
            ClientSecrets secrets = new ClientSecrets()
            {
                ClientId = googleClientId,
                ClientSecret = googleClientSecret,
            };

            return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, "user", CancellationToken.None).Result;
        }

        string Base64UrlEncode(string input)
        {
            var data = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }

        private void SendEmails(List<(string user, string server)> users_to_send_email, bool isTaskTakesTime)
        {
            if (service_email == null || users_to_send_email == null || users_to_send_email.Count == 0) return;

            foreach (var usr in users_to_send_email)
            {
                SendEmail(usr.user, usr.server, isTaskTakesTime);
            }
        }

        private void SendEmail(string user, string server, bool isTakenTime)
        {
            if (service_email == null) return;

            string sender = user + "@fmp-berlin.de";

            try
            {
                #region first char to capital letter
                string[] cols_user = Regex.Split(user, "\\.");
                string user_capital = char.ToUpper(cols_user[0][0]) + cols_user[0].Substring(1);
                if (cols_user.Length > 1)
                    user_capital = char.ToUpper(cols_user[1][0]) + cols_user[1].Substring(1) + " " + char.ToUpper(cols_user[0][0]) + cols_user[0].Substring(1);
                #endregion
                string message = "";

                if (isTakenTime)
                    message = $"To: {sender}\r\nSubject: WARNING: Check the {server} server\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<p>Dear {user_capital},<br/><br/>You have started a task on the {server} server and it is taking a long time to finish. Verify that this task completed. If so, discard this message and update the status on <u><i>'Remote Desktop Connection'</i></u> app.<br/><br/>Best regards,<br/>AG Fliu</p>";
                else
                    message = $"To: {sender}\r\nSubject: WARNING: Connection on {server} server\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<p>Dear {user_capital},<br/><br/>You have been disconnected from the {server} server because you have been connected for more than 3 hours.<br/>Please check if your tasks have been finished. If so, update the status on <u><i>'Remote Desktop Connection'</i></u> app.<br/><br/>Best regards,<br/>AG Fliu</p>";
                var msg = new Google.Apis.Gmail.v1.Data.Message();
                msg.Raw = Base64UrlEncode(message.ToString());
                service_email.Users.Messages.Send(msg, "me").Execute();

                //Update time in google sheet
                var data = ReadSheetUsers(server);
                int current_user_index = data.FindIndex(a => a.Name.Equals(user));
                UpdateInfo(server, current_user_index, false, false, true);
            }
            catch (Exception)
            {
            }
        }

        private void Init()
        {
            string[] Scopes = { GmailService.Scope.GmailSend, SheetsService.Scope.Spreadsheets };
            string googleClientID = "680087468736-skuel9mfejmals227c7kqrge7nm89qnf.apps.googleusercontent.com";
            string googleClientSecret = "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";

            try
            {
                UserCredential credential = Login(googleClientID, googleClientSecret, Scopes);

                service_email = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                service_sheets = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
            catch (Exception)
            {
            }

        }

        private void ReadGoogleSheets()
        {
            data_agms2_users = ReadSheetUsers("AGMS2");
            data_agms2_users.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
            data_agms2_software = ReadSheetSoftware("Installed_software_AGMS2");
            data_agms2_software.Sort((a, b) => a.Name.CompareTo(b.Name));

            data_agms3_users = ReadSheetUsers("AGMS3");
            data_agms3_users.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
            data_agms3_software = ReadSheetSoftware("Installed_software_AGMS3");
            data_agms3_software.Sort((a, b) => a.Name.CompareTo(b.Name));
        }

        private List<(string user, string server)> CheckLongerUsers()
        {
            List<(string user, string server)> user_to_send_email = new();
            DateTime currentDate = DateTime.Now;

            #region AGMS2
            List<DataGridItemUser> userslogged = data_agms2_users.Where(a => a.IsLogged).ToList();

            foreach (var dgi in userslogged)
            {
                DateTime user_date = DateTime.ParseExact(dgi.Date, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.Days > 0 || difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                    user_to_send_email.Add((dgi.Name, "AGMS2"));
            }
            #endregion

            #region AGMS3
            userslogged = data_agms3_users.Where(a => a.IsLogged).ToList();

            foreach (var dgi in userslogged)
            {
                DateTime user_date = DateTime.ParseExact(dgi.Date, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.Days > 0 || difference.Hours > TIMEOUT_HOURS_USER_CONNECTION)
                    user_to_send_email.Add((dgi.Name, "AGMS3"));
            }
            #endregion

            return user_to_send_email.Distinct().ToList();
        }

        private List<(string user, string server)> CheckLongerTasks()
        {
            List<(string user, string server)> user_to_send_email = new();
            DateTime currentDate = DateTime.Now;

            #region AGMS2
            List<DataGridItemUser> tasksNotFinished = data_agms2_users.Where(a => !a.HasTaskFinished).ToList();

            foreach (var dgi in tasksNotFinished)
            {
                DateTime user_date = DateTime.ParseExact(dgi.TaskDate, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.TotalDays > TIMEOUT_DAYS_USER_TASKS)
                    user_to_send_email.Add((dgi.Name, "AGMS2"));
            }
            #endregion

            #region AGMS3
            tasksNotFinished = data_agms3_users.Where(a => !a.HasTaskFinished).ToList();

            foreach (var dgi in tasksNotFinished)
            {
                DateTime user_date = DateTime.ParseExact(dgi.TaskDate, "dd/MM/yyyy HH:mm:ss", cultureInfo);
                TimeSpan difference = currentDate - user_date;
                if (difference.TotalDays > TIMEOUT_DAYS_USER_TASKS)
                    user_to_send_email.Add((dgi.Name, "AGMS3"));
            }
            #endregion

            return user_to_send_email.Distinct().ToList();
        }

        private void LoadDatagrid()
        {
            if (data_agms2_users == null || data_agms3_users == null) return;
            DataGridAGMS2.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS2.ItemsSource = null;
                data_agms2_users.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
                DataGridAGMS2.ItemsSource = data_agms2_users;
            }));

            DataGridAGMS3.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS3.ItemsSource = null;
                data_agms3_users.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
                DataGridAGMS3.ItemsSource = data_agms3_users;
            }));

            if (data_agms2_software == null || data_agms3_software == null) return;
            DataGridAGMS2_Software.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS2_Software.ItemsSource = null;
                DataGridAGMS2_Software.ItemsSource = data_agms2_software;
            }));

            DataGridAGMS3_Software.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS3_Software.ItemsSource = null;
                DataGridAGMS3_Software.ItemsSource = data_agms3_software;
            }));
        }
        private void ConnectAGMS(string ipAddress, string username, string password, out string error)
        {
            try
            {
                var process = new Process();
                process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdKey.exe");
                process.StartInfo.Arguments = String.Format(@"/generic:{0} /user:{1} /pass:{2}", ipAddress, username, password);
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();

                connectServerProcess = new Process();
                connectServerProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                connectServerProcess.StartInfo.Arguments = String.Format(@"/f /v:{0}", ipAddress);
                connectServerProcess.EnableRaisingEvents = true;
                connectServerProcess.StartInfo.UseShellExecute = false;
                connectServerProcess.StartInfo.RedirectStandardOutput = true;
                connectServerProcess.Start();
                connectServerProcess.PriorityClass = ProcessPriorityClass.High;
                connectServerProcess.PriorityBoostEnabled = true;
                isConnected = true;

                dispatcherTimer_checkProcessAlive.Start();
                connected_initial_time = connectServerProcess.StartTime;

                var outputResultPromise = connectServerProcess.StandardOutput.ReadToEndAsync();
                outputResultPromise.ContinueWith(o => isConnected = false);
                error = "";
            }
            catch (Exception e)
            {
                error = "";
            }
        }

        private void ButtonConnect_Click(object sender, RoutedEventArgs e)
        {
            dispatcherTimer.Stop();
            string ipAddress = "";
            selected_server = -1;
            if (AGMS_tab.SelectedIndex == 0)//AGMS2
            {
                ipAddress = "10.10.65.31";
                selected_server = 2;
            }
            else if (AGMS_tab.SelectedIndex == 1)//AGMS3
            {
                ipAddress = "10.10.65.49";
                selected_server = 3;
            }

            DataGridItemUser dgi;
            if (!CanConnect(selected_server, out dgi))
            {

                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));

                //Update google sheets
                UpdateSheet(selected_server, true, false);

                string username = "sysfliu";
                string password = "Agms2023!";
                ConnectAGMS(ipAddress, username, password, out error_taken_time_connected);
            }
            else
            {
                System.Windows.MessageBox.Show(
                        "The following user is connected on AGMS" + selected_server + ": " + (dgi != null ? dgi.Name : "---") + "!\nPlease contact him/her to log out the server or try later!",
                        "The Liu Lab :: Information",
                        (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                        (System.Windows.MessageBoxImage)MessageBoxIcon.Warning);
                dispatcherTimer.Start();
            }
        }

        private bool CanConnect(int selected_server, out DataGridItemUser dgi)
        {
            DateTime currentTime = DateTime.Now;
            TimeSpan difference = currentTime - last_refresh;
            if (difference.TotalSeconds < INTERVAL_TIME_REFRESH_SECONDS)
            {
                ReadSheets();
                last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }

            while (data_agms2_users.Count == 0 || data_agms3_users.Count == 0)
            {
                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));
                System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                ReadSheets();

                last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }
            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
            LoadDatagrid();

            if (selected_server == 2 || selected_server == 1)
            {
                dgi = data_agms2_users.Where(a => a.IsLogged).FirstOrDefault();
                if (dgi != null)
                    return true;
                else
                    return false;
            }
            else if (selected_server == 3)
            {
                dgi = data_agms3_users.Where(a => a.IsLogged).FirstOrDefault();
                if (dgi != null)
                    return true;
                else
                    return false;
            }

            dgi = null;
            return true;
        }

        private void UpdateSheet(int selected_server, bool isLogged, bool hasTaskFinished)
        {
            string sheet = "AGMS" + selected_server;
            try
            {
                if (selected_server == 2 || selected_server == 1)
                {
                    data_agms2_users = ReadSheetUsers(sheet);
                    while (data_agms2_users.Count == 0)
                    {
                        System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                        data_agms2_users = ReadSheetUsers(sheet);
                    }

                    int current_user_index = data_agms2_users.FindIndex(a => a.Name.Equals(TBUser.Text));
                    if (current_user_index == -1)//Add information
                    {
                        AddInfo(sheet);
                    }
                    else
                    {
                        //if task status is true (has been finished), then date can be updated at the moment of login
                        var current_user_previous_task_status = data_agms2_users[current_user_index].HasTaskFinished;
                        bool update_date = current_user_previous_task_status;

                        UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, update_date);
                        var update_data = data_agms2_users[current_user_index];
                        update_data.IsLogged = isLogged;
                        update_data.HasTaskFinished = hasTaskFinished;
                    }
                }
                else if (selected_server == 3)
                {
                    data_agms3_users = ReadSheetUsers(sheet);
                    while (data_agms3_users.Count == 0)
                    {
                        System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                        data_agms3_users = ReadSheetUsers(sheet);
                    }

                    int current_user_index = data_agms3_users.FindIndex(a => a.Name.Equals(TBUser.Text));
                    if (current_user_index == -1)//Add information
                    {
                        AddInfo(sheet);
                    }
                    else
                    {
                        //if task status is true (has been finished), then date can be updated at the moment of login
                        var current_user_previous_task_status = data_agms3_users[current_user_index].HasTaskFinished;
                        bool update_date = current_user_previous_task_status;

                        UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, update_date);
                        var update_data = data_agms3_users[current_user_index];
                        update_data.IsLogged = isLogged;
                        update_data.HasTaskFinished = hasTaskFinished;
                    }
                }
            }
            catch (Exception e)
            {
            }
        }

        private void UpdateInfo(string sheet, int current_user_index, bool isLogged, bool hasTaskFinished, bool updateDate)
        {
            string range = "";
            var valueRange = new ValueRange();
            List<object> oblist;
            if (!hasTaskFinished && !updateDate)
            {
                // Specifying Column Range for reading...
                range = $"{sheet}!C" + (current_user_index + 2) + ":E" + (current_user_index + 2);
                oblist = new List<object>() { DateTime.Now, isLogged, hasTaskFinished };
            }
            else
            {
                // Specifying Column Range for reading...
                range = $"{sheet}!B" + (current_user_index + 2) + ":E" + (current_user_index + 2);
                oblist = new List<object>() { DateTime.Now, DateTime.Now, isLogged, hasTaskFinished };
            }
            valueRange.Values = new List<IList<object>> { oblist };

            // Performing Update Operation...
            try
            {
                var updateRequest = service_sheets.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = updateRequest.Execute();
            }
            catch (Exception e)
            {
                System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, updateDate);
            }
        }

        private void AddInfo(string sheet)
        {
            // Specifying Column Range for reading...
            var range = $"{sheet}!A:E";
            var valueRange = new ValueRange();
            var oblist = new List<object>() { TBUser.Text, DateTime.Now, DateTime.Now, true, false };
            valueRange.Values = new List<IList<object>> { oblist };
            try
            {
                // Append the above record...
                var appendRequest = service_sheets.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = appendRequest.Execute();
            }
            catch (Exception e)
            {
                System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                AddInfo(sheet);
            }
        }

        private List<DataGridItemUser> ReadSheetUsers(string sheet)
        {
            List<DataGridItemUser> _data = new();
            try
            {
                // Specifying Column Range for reading...
                var range = $"{sheet}!A:E";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service_sheets.Spreadsheets.Values.Get(SpreadsheetId, range);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                // Executing Read Operation...
                var response = request.Execute();
                // Getting all records from Column A to E...
                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row[0].ToString().StartsWith("Username")) continue;
                        //Username, Date, Are you logged, Has the task finished
                        DateTime taskDate = Convert.ToDateTime(row[1].ToString());
                        string taskDate_str = taskDate.ToString("dd") + "/" + taskDate.ToString("MM") + "/" + taskDate.ToString("yyyy") + " " + taskDate.ToString("HH:mm:ss");
                        DateTime date = Convert.ToDateTime(row[2].ToString());
                        string date_str = date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.ToString("yyyy") + " " + date.ToString("HH:mm:ss");
                        var row_data = new DataGridItemUser(sheet, row[0].ToString(), date_str, date, taskDate_str, taskDate, Convert.ToBoolean(row[3]), Convert.ToBoolean(row[4]));
                        _data.Add(row_data);
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
            }
            catch (Exception e)
            {
            }

            return _data;
        }

        private void ReadSheets()
        {
            List<DataGridItemUser> _data_users = new();
            List<DataGridItemSoftware> _data_softwares = new();
            try
            {
                var ranges = new List<string> {
                    "AGMS2!A:E",
                    "AGMS3!A:E",
                    "Installed_software_AGMS2!A:B",
                    "Installed_software_AGMS3!A:B"
                };
                var request = new SpreadsheetsResource.ValuesResource.BatchGetRequest(service_sheets, SpreadsheetId);
                request.Ranges = ranges;
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                // Executing Read Operation...
                var response = request.Execute();

                foreach (var valueRange in response.ValueRanges)
                {
                    var values = valueRange.Values;

                    if (values != null && values.Count > 0)
                    {
                        if (values[0][0] != null && values[0][0].ToString().StartsWith("Username"))
                        {
                            var range = Regex.Split(valueRange.Range, "!")[0];

                            for (int i = 1; i < values.Count; i++)
                            {
                                var row = values[i];
                                //Username, Date, Are you logged, Has the task finished
                                DateTime taskDate = Convert.ToDateTime(row[1].ToString());
                                string taskDate_str = taskDate.ToString("dd") + "/" + taskDate.ToString("MM") + "/" + taskDate.ToString("yyyy") + " " + taskDate.ToString("HH:mm:ss");
                                DateTime date = Convert.ToDateTime(row[2].ToString());
                                string date_str = date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.ToString("yyyy") + " " + date.ToString("HH:mm:ss");
                                var row_data = new DataGridItemUser(range, row[0].ToString(), date_str, date, taskDate_str, taskDate, Convert.ToBoolean(row[3]), Convert.ToBoolean(row[4]));
                                _data_users.Add(row_data);
                            }
                        }
                        else if (values[0][0] != null && values[0][0].ToString().StartsWith("Software"))
                        {
                            var range = Regex.Split(valueRange.Range, "_")[2];
                            range = Regex.Split(range, "!")[0];

                            for (int i = 1; i < values.Count; i++)
                            {
                                var row = values[i];
                                //Username, Date, Are you logged, Has the task finished
                                var row_data = new DataGridItemSoftware(range, row[0].ToString(), row[1].ToString());
                                _data_softwares.Add(row_data);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
            }

            data_agms2_users = _data_users.Where(a => a.Server == "AGMS2").ToList();
            data_agms2_software = _data_softwares.Where(a => a.Server == "AGMS2").ToList();
            data_agms3_users = _data_users.Where(a => a.Server == "AGMS3").ToList();
            data_agms3_software = _data_softwares.Where(a => a.Server == "AGMS3").ToList();
        }
        private List<DataGridItemSoftware> ReadSheetSoftware(string sheet)
        {
            List<DataGridItemSoftware> _data = new();
            try
            {
                // Specifying Column Range for reading...
                var range = $"{sheet}!A:B";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service_sheets.Spreadsheets.Values.Get(SpreadsheetId, range);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                // Ecexuting Read Operation...
                var response = request.Execute();
                // Getting all records from Column A to B...
                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row[0].ToString().StartsWith("Software")) continue;
                        //Username, Date, Are you logged, Has the task finished
                        var row_data = new DataGridItemSoftware(sheet, row[0].ToString(), row[1].ToString());
                        _data.Add(row_data);
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
            }
            catch (Exception e)
            {
            }

            return _data;
        }

        private void Process_Exited(object? sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(error_taken_time_connected) && error_taken_time_connected.Equals("closed"))//Check if user took time to close RDP
                return;

            selected_server = -1;

            AGMS_tab.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                if (AGMS_tab.SelectedIndex == 0)//AGMS2
                    selected_server = 2;
                else if (AGMS_tab.SelectedIndex == 1)//AGMS3
                    selected_server = 3;

                var response = System.Windows.Forms.MessageBox.Show($"Has your task been finished?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (response == System.Windows.Forms.DialogResult.Yes)
                {
                    //Update google sheets
                    UpdateSheet(selected_server, false, true);
                }
                else
                {
                    //Update google sheets
                    UpdateSheet(selected_server, false, false);
                }

                LoadDatagrid();

                dispatcherTimer.Start();

                System.Windows.MessageBox.Show(
                            "The server has been closed successfully!",
                            "The Liu Lab :: Information",
                            (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                            (System.Windows.MessageBoxImage)MessageBoxIcon.Information);
            }));

            isConnected = false;
        }

        private void DataGridAGMS2_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void DataGridAGMS3_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }
    }
}
