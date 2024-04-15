using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Security;
using System.Reflection.PortableExecutable;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static Google.Apis.Gmail.v1.UsersResource.SettingsResource;
using System.Windows.Documents;
using System.Windows.Forms;
using static RemoteDesktopConnection.MainWindow;
using RemoteDesktopConnection.Model;

namespace RemoteDesktopConnection.Control.Database
{
    public class Connection
    {
        private static readonly string ApplicationName = "Google Sheet API .NET Quickstart";
        public static string SpreadsheetId;
        private static UserCredential UserCredential;

        public static SheetsService service_sheets;
        public static GmailService service_email;
        private CultureInfo cultureInfo = new CultureInfo("en-US");

        public static double Refresh_time = 0;

        private const int OFFSET_TIME_REFRESH_MILLISECONDS = 1000;

        public Connection()
        {
        }

        public static void Init()
        {
            string[] Scopes = {
                GmailService.Scope.GmailSend,
                SheetsService.Scope.Spreadsheets };

            Management.SetDatabase();

            RemoteDesktopConnection.Model.Database db = Management.GetDatabase();
            if (db == null) return;

            SpreadsheetId = db.SpreadsheetID;

            try
            {
                UserCredential = Login(db.GoogleClientID, db.GoogleClientSecret, Scopes);

                service_email = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = UserCredential,
                    ApplicationName = ApplicationName
                });

                service_sheets = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = UserCredential,
                    ApplicationName = ApplicationName
                });
            }
            catch (Exception)
            {
            }
        }
        public async static void RevokeConnection()
        {
            try
            {
                if (UserCredential != null)
                {
                    await UserCredential.RevokeTokenAsync(CancellationToken.None);
                }
            }
            catch (Exception)
            {
            }
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

        public static void ReadSheets()
        {
            List<Model.User> _users = new();
            List<Model.Software> _softwares = new();
            List<Model.Server> _servers = new();

            try
            {

                var ranges = new List<string> {
                    "AGMS2!A:F",
                    "AGMS3!A:F",
                    "AGMS4!A:F",
                    "Installed_software_AGMS2!A:B",
                    "Installed_software_AGMS3!A:B",
                    "Installed_software_AGMS4!A:B",
                    "Maintenance!A:C",
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
                                var row_data = new Model.User(range, row[0].ToString(), date_str, taskDate_str, date, taskDate, Convert.ToBoolean(row[3]), Convert.ToBoolean(row[4]), Convert.ToString(row[5]));
                                _users.Add(row_data);
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
                                var row_data = new Model.Software(range, row[0].ToString(), row[1].ToString());
                                _softwares.Add(row_data);
                            }
                        }
                        else if (values[0][0] != null && values[0][0].ToString().StartsWith("Server"))
                        {
                            var range = Regex.Split(valueRange.Range, "!")[0];

                            for (int i = 1; i < values.Count; i++)
                            {
                                var row = values[i];
                                var row_data = new Model.Server(row[0].ToString(), Convert.ToBoolean(row[1]), row[2].ToString());
                                _servers.Add(row_data);
                            }
                        }

                    }
                }
            }
            catch (System.Threading.Tasks.TaskCanceledException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (Google.GoogleApiException exc)
            {
                if (!(exc.Error.Code == 500 && exc.Error.Message.Contains("The service sheets has thrown an exception")))
                    throw new Exception("reset_database");
            }
            catch (System.FormatException e)
            {
                Console.WriteLine(e.Message);
                if (e.Message.Contains("was not recognized as a valid"))
                {
                    System.Windows.MessageBox.Show(
                                            "The database is not recognized.\nPlease contact the administrator.",
                                            "RemoteDesktopConnection :: Error",
                                            (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                                            (System.Windows.MessageBoxImage)MessageBoxIcon.Error);

                    Connection.RevokeConnection();
                    System.Environment.Exit(1);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Management.Users_agms2 = new();
            Management.Users_agms3 = new();
            Management.Users_agms4 = new();
            Management.Software_agms2 = new();
            Management.Software_agms3 = new();
            Management.Software_agms4 = new();
            Management.Servers = new();

            if (_servers.Count > 0)
            {
                Management.Servers = _servers;
            }
            if (_users.Count > 0)
            {
                Management.Users_agms2 = _users.Where(a => a.Server == "AGMS2").ToList();
                Management.Users_agms3 = _users.Where(a => a.Server == "AGMS3").ToList();
                Management.Users_agms4 = _users.Where(a => a.Server == "AGMS4").ToList();
            }
            if (_softwares.Count > 0)
            {
                Management.Software_agms2 = _softwares.Where(a => a.Server == "AGMS2").ToList();
                Management.Software_agms3 = _softwares.Where(a => a.Server == "AGMS3").ToList();
                Management.Software_agms4 = _softwares.Where(a => a.Server == "AGMS4").ToList();
            }

        }

        private static List<Model.User> ReadSheetUsers(string sheet)
        {
            List<Model.User> _data = new();
            try
            {
                // Specifying Column Range for reading...
                var range = $"{sheet}!A:F";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service_sheets.Spreadsheets.Values.Get(SpreadsheetId, range);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                // Executing Read Operation...
                var response = request.Execute();
                // Getting all records from Column A to F...
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
                        var row_data = new Model.User(sheet, row[0].ToString(), date_str, taskDate_str, date, taskDate, Convert.ToBoolean(row[3]), Convert.ToBoolean(row[4]), Convert.ToString(row[5]));
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

        private static List<Model.Software> ReadSheetSoftware(string sheet)
        {
            List<Model.Software> _data = new();
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
                        var row_data = new Model.Software(sheet, row[0].ToString(), row[1].ToString());
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

        public static void SendEmails(List<(string user, string email, string server)> users_to_send_email, bool isTaskTakesTime)
        {
            if (service_email == null || users_to_send_email == null || users_to_send_email.Count == 0) return;

            foreach (var usr in users_to_send_email)
            {
                SendEmail(usr.user, usr.email, usr.server, isTaskTakesTime);
            }
        }

        public static void SendEmail(string user, string email, string server, bool isTakenTime)
        {
            if (service_email == null) return;

            string sender = email;

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
                    message = $"To: {sender}\r\nSubject: WARNING: Check the {server} server\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<p>Dear {user_capital},<br/><br/>You have started a task on the {server} server and it is taking a long time to finish. Verify that this task completed. If so, discard this message and update the status on <u><i>'Remote Desktop Connection'</i></u> app.<br/><br/>Best regards,<br/>AG FLiu</p>";
                else if (!String.IsNullOrEmpty(server))
                    message = $"To: {sender}\r\nSubject: WARNING: Connection on {server} server\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<p>Dear {user_capital},<br/><br/>You have been disconnected from the {server} server because you have been connected for more than 1 hour.<br/>Please check if your tasks have been finished. If so, update the status on <u><i>'Remote Desktop Connection'</i></u> app.<br/><br/>Best regards,<br/>AG FLiu</p>";
                else
                    message = $"To: {sender}\r\nSubject: WARNING: Remote Deskop Connection was closed\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<p>Dear {user_capital},<br/><br/>The Remote Desktop Connection app was closed because you were inactive for more than 5 minutes.<br/><br/>Best regards,<br/>AG FLiu</p>";
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

        private static void UpdateInfo(string sheet, int current_user_index, bool isLogged, bool hasTaskFinished, bool updateDate)
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

        public static void UpdateSheet(int selected_server, bool isLogged, bool hasTaskFinished)
        {
            string sheet = "AGMS" + selected_server;
            try
            {
                if (selected_server == 2 || selected_server == 1)
                {
                    Management.Users_agms2 = ReadSheetUsers(sheet);
                    while (Management.Users_agms2.Count == 0)
                    {
                        System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                        Management.Users_agms2 = ReadSheetUsers(sheet);
                    }

                    int current_user_index = Management.Users_agms2.FindIndex(a => a.Name.Equals(Management.GetUser()));
                    if (current_user_index == -1)//Add information
                    {
                        AddInfo(sheet);
                    }
                    else
                    {
                        //if task status is true (has been finished), then date can be updated at the moment of login
                        var current_user_previous_task_status = Management.Users_agms2[current_user_index].HasTaskFinished;
                        bool update_date = current_user_previous_task_status;

                        UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, update_date);
                        var update_data = Management.Users_agms2[current_user_index];
                        update_data.IsLogged = isLogged;
                        update_data.HasTaskFinished = hasTaskFinished;
                    }
                }
                else if (selected_server == 3)
                {
                    Management.Users_agms3 = ReadSheetUsers(sheet);
                    while (Management.Users_agms3.Count == 0)
                    {
                        System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                        Management.Users_agms3 = ReadSheetUsers(sheet);
                    }

                    int current_user_index = Management.Users_agms3.FindIndex(a => a.Name.Equals(Management.GetUser()));
                    if (current_user_index == -1)//Add information
                    {
                        AddInfo(sheet);
                    }
                    else
                    {
                        //if task status is true (has been finished), then date can be updated at the moment of login
                        var current_user_previous_task_status = Management.Users_agms3[current_user_index].HasTaskFinished;
                        bool update_date = current_user_previous_task_status;

                        UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, update_date);
                        var update_data = Management.Users_agms3[current_user_index];
                        update_data.IsLogged = isLogged;
                        update_data.HasTaskFinished = hasTaskFinished;
                    }
                }
                else if (selected_server == 4)
                {
                    Management.Users_agms4 = ReadSheetUsers(sheet);
                    while (Management.Users_agms4.Count == 0)
                    {
                        System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                        Management.Users_agms4 = ReadSheetUsers(sheet);
                    }

                    int current_user_index = Management.Users_agms4.FindIndex(a => a.Name.Equals(Management.GetUser()));
                    if (current_user_index == -1)//Add information
                    {
                        AddInfo(sheet);
                    }
                    else
                    {
                        //if task status is true (has been finished), then date can be updated at the moment of login
                        var current_user_previous_task_status = Management.Users_agms4[current_user_index].HasTaskFinished;
                        bool update_date = current_user_previous_task_status;

                        UpdateInfo(sheet, current_user_index, isLogged, hasTaskFinished, update_date);
                        var update_data = Management.Users_agms4[current_user_index];
                        update_data.IsLogged = isLogged;
                        update_data.HasTaskFinished = hasTaskFinished;
                    }
                }
            }
            catch (Exception e)
            {
            }
        }

        private static void AddInfo(string sheet)
        {
            // Specifying Column Range for reading...
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();
            var oblist = new List<object>() { Management.GetUser(), DateTime.Now, DateTime.Now, true, false, (Management.GetUser()) + "@fmp-berlin.de" };
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

        private static string Base64UrlEncode(string input)
        {
            var data = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }

    }
}
