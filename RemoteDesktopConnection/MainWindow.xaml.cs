using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Logging;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic.ApplicationServices;
using RemoteDesktopConnection.Control;
using RemoteDesktopConnection.Control.Database;
using RemoteDesktopConnection.Model;
using RemoteDesktopConnection.Util;
using RemoteDesktopConnection.Viewer;
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
using System.Runtime.InteropServices;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace RemoteDesktopConnection
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        private System.Windows.Threading.DispatcherTimer dispatcherProgressBar = new System.Windows.Threading.DispatcherTimer();
        private System.Windows.Threading.DispatcherTimer dispatcherTimer_checkProcessAlive = new System.Windows.Threading.DispatcherTimer();
        private System.Windows.Threading.DispatcherTimer dispatcherTimer_checkCloseSoftware = new System.Windows.Threading.DispatcherTimer();
        private double progressbar_time = 0;
        private readonly string ApplicationName = "Google Sheet API .NET Quickstart";

        private int timeUntilCloseSoftware { get; set; }
        private string error_taken_time_connected = "";
        private string current_user = "";

        private int selected_server = -1;


        private const int OFFSET_TIME_REFRESH_MILLISECONDS = 1000;

        [STAThread]
        private void ConnectDatabase()
        {
            Connection.Init();
            Connection.ReadSheets();
        }
        private async void StartApp()
        {
            var wait_screen = Util.Util.CallWaitWindow("Welcome to Remote Desktop Connection", "Please wait, we are loading data...");
            MainGrid.Children.Add(wait_screen);
            Grid.SetRowSpan(wait_screen, 4);
            Grid.SetRow(wait_screen, 0);
            var rows = MainGrid.RowDefinitions;
            rows[2].Height = new GridLength(2, GridUnitType.Star);
            await Task.Run(() => ConnectDatabase());

            MainGrid.Children.Remove(wait_screen);
            rows[2].Height = GridLength.Auto;

            current_user = Control.Management.GetUser();
            TBUser.Text = current_user;

            if (Management.Users_agms2.Count > 0 && Management.Users_agms3.Count > 0)
            {
                Management.last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = Management.last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }

            this.Closed += new EventHandler(Window_Closed);

            LoadDatagrid();
            Connection.SendEmails(Management.CheckLongerTasks(), true);
            Connection.SendEmails(Management.CheckLongerUsers(), false);

            dispatcherProgressBar.Tick += new EventHandler(dispatcherProgressBar_Tick);
            dispatcherProgressBar.Interval = new TimeSpan(0, 0, 1);
            dispatcherProgressBar.Start();

            dispatcherTimer.Tick += new EventHandler(dispatcherReadSpreadsheet_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, Management.INTERVAL_TIME_REFRESH_SECONDS);
            dispatcherTimer.Start();

            dispatcherTimer_checkCloseSoftware.Tick += new EventHandler(dispatcherTimer_checkCloseSoftware_Tick);
            dispatcherTimer_checkCloseSoftware.Interval = new TimeSpan(0, 0, Management.INTERVAL_TIME_REFRESH_SECONDS);
            dispatcherTimer_checkCloseSoftware.Start();

            dispatcherTimer_checkProcessAlive.Tick += new EventHandler(dispatcherTimer_checkProcessAlive_Tick);
            dispatcherTimer_checkProcessAlive.Interval = new TimeSpan(0, 0, 1);
        }
        public MainWindow()
        {
            InitializeComponent();
            DateTime dt = DateTime.Now;
            AddHyperlink(InfoLabLabel, $"The Liu Lab @ {dt.Year} - v. {Util.Util.GetVersion()} - All rights reserved!");
            StartApp();
        }
        private void AddHyperlink(TextBlock textBlock, string processing_time)
        {
            textBlock.Inlines.Clear();

            // Create a new Hyperlink
            Hyperlink hyperlink = new Hyperlink();
            hyperlink.Inlines.Add("Diogo Borges Lima");
            hyperlink.NavigateUri = new System.Uri("https://diogobor.droppages.com/");
            hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
            // Add the Hyperlink to the TextBlock
            textBlock.Inlines.Add(processing_time);
            textBlock.Inlines.Add(" Developed by ");
            textBlock.Inlines.Add(hyperlink);
            textBlock.Inlines.Add(".");
        }
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            // Open the link in the default web browser
            try
            {
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri.ToString(),
                    UseShellExecute = true,
                });
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show(
                        "There is no internet connection.",
                        "Warning",
                        (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                        (System.Windows.MessageBoxImage)MessageBoxIcon.Information);
                throw;
            }
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            if (Management.ConnectServerProcess != null && !Management.ConnectServerProcess.HasExited)
            {
                dispatcherTimer_checkProcessAlive.Stop();
                error_taken_time_connected = "closed";

                Management.ConnectServerProcess.Kill();
                if (selected_server == -1) return;

                var response = System.Windows.Forms.MessageBox.Show($"Has your task been finished?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (response == System.Windows.Forms.DialogResult.Yes)
                {
                    //Update google sheets
                    Connection.UpdateSheet(selected_server, false, true);
                }
                else
                {
                    //Update google sheets
                    Connection.UpdateSheet(selected_server, false, false);
                }

                System.Windows.MessageBox.Show(
                            "The server has been closed successfully!",
                            "The Liu Lab :: Information",
                            (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                            (System.Windows.MessageBoxImage)MessageBoxIcon.Information);
            }
        }

        private void dispatcherProgressBar_Tick(object sender, EventArgs e)
        {
            if (progressbar_time < Management.INTERVAL_TIME_REFRESH_SECONDS)
                progressbar_time++;
            else
                progressbar_time = 0;
            ProgressBarRefresh.Value = progressbar_time * (100.0 / Management.INTERVAL_TIME_REFRESH_SECONDS);
        }
        private void dispatcherReadSpreadsheet_Tick(object sender, EventArgs e)
        {
            do
            {
                Connection.ReadSheets();
                if (Management.Users_agms2.Count > 0 && Management.Users_agms3.Count > 0 && Management.Users_agms4.Count > 0)
                {
                    Management.last_refresh = DateTime.Now;
                    LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = Management.last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
                }

                if (Management.Users_agms2.Count == 0 || Management.Users_agms3.Count == 0 || Management.Users_agms4.Count == 0)
                {
                    ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));
                    System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                }
            }
            while (Management.Users_agms2.Count == 0 || Management.Users_agms3.Count == 0 || Management.Users_agms4.Count == 0);
            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));

            LoadDatagrid();
        }
        private void dispatcherTimer_checkProcessAlive_Tick(object sender, EventArgs e)
        {
            bool isAlive = Management.CheckProcessStatus();

            if (!Management.IsConnected || !isAlive)
            {
                Process_Exited(null, null);
                dispatcherTimer_checkProcessAlive.Stop();
                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
            }
            else
            {
                TimeSpan difference = DateTime.Now - Management.Connected_initial_time;

                if (difference.Hours > Management.TIMEOUT_HOURS_USER_CONNECTION)
                {
                    Management.ConnectServerProcess.Kill();

                    //Send an email to the user
                    TBUser.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
                    {
                        string email = string.Empty;
                        if (selected_server == 2)
                            email = Management.Users_agms2.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                        else if (selected_server == 3)
                            email = Management.Users_agms3.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                        else
                            email = Management.Users_agms4.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                        Connection.SendEmail(TBUser.Text, email, ("AGMS" + selected_server), false);
                        dispatcherTimer.Start();
                        progressbar_time = 0;
                    }));
                    dispatcherTimer_checkProcessAlive.Stop();
                    Management.CurrentPID = -1;
                }
            }
        }
        private void dispatcherTimer_checkCloseSoftware_Tick(object sender, EventArgs e)
        {
            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                if (ButtonConnect.IsEnabled == true)
                {
                    if (timeUntilCloseSoftware > 2)
                    {
                        dispatcherProgressBar.Stop();
                        dispatcherTimer_checkCloseSoftware.Stop();
                        dispatcherTimer_checkProcessAlive.Stop();
                        dispatcherTimer.Stop();
                        progressbar_time = int.MinValue;
                        timeUntilCloseSoftware = 0;

                        IdleWindow iw = new IdleWindow();
                        iw.ShowDialog();

                        if (iw.CloseSoftware == true)
                        {
                            TBUser.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
                            {
                                string email = string.Empty;
                                if (selected_server == 2)
                                    email = Management.Users_agms2.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                                else if (selected_server == 3)
                                    email = Management.Users_agms3.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                                else
                                    email = Management.Users_agms4.Where(a => a.Name == TBUser.Text).FirstOrDefault().Email;
                                Connection.SendEmail(TBUser.Text, email, "", false);
                                progressbar_time = 0;
                            }));

                            this.Close();
                        }
                        else
                        {
                            progressbar_time = 0;
                            dispatcherProgressBar.Start();
                            dispatcherTimer_checkCloseSoftware.Start();
                            dispatcherTimer.Start();
                        }
                    }
                    else
                        timeUntilCloseSoftware++;
                }

            }));

        }

        private void LoadDatagrid()
        {
            if (Control.Management.Users_agms2 == null || Control.Management.Users_agms3 == null || Control.Management.Users_agms4 == null) return;
            DataGridAGMS2.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS2.ItemsSource = null;
                Control.Management.Users_agms2.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
                List<Model.User> users = Control.Management.Users_agms2;
                Parallel.ForEach(users, a =>
                {
                    a._isLogged = a.IsLogged ? "✓" : "-";
                    a._hasTaskFinished = a.HasTaskFinished ? "✓" : "-";
                });
                DataGridAGMS2.ItemsSource = users;
            }));

            DataGridAGMS3.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS3.ItemsSource = null;
                Control.Management.Users_agms3.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
                List<Model.User> users = Control.Management.Users_agms3;
                Parallel.ForEach(users, a =>
                {
                    a._isLogged = a.IsLogged ? "✓" : "-";
                    a._hasTaskFinished = a.HasTaskFinished ? "✓" : "-";
                });
                DataGridAGMS3.ItemsSource = users;
            }));

            DataGridAGMS4.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS4.ItemsSource = null;
                Control.Management.Users_agms4.Sort((a, b) => b.DateOriginal.CompareTo(a.DateOriginal));
                List<Model.User> users = Control.Management.Users_agms4;
                Parallel.ForEach(users, a =>
                {
                    a._isLogged = a.IsLogged ? "✓" : "-";
                    a._hasTaskFinished = a.HasTaskFinished ? "✓" : "-";
                });
                DataGridAGMS4.ItemsSource = users;
            }));

            if (Control.Management.Software_agms2 == null || Control.Management.Software_agms3 == null) return;
            DataGridAGMS2_Software.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS2_Software.ItemsSource = null;
                DataGridAGMS2_Software.ItemsSource = Control.Management.Software_agms2;
            }));

            DataGridAGMS3_Software.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS3_Software.ItemsSource = null;
                DataGridAGMS3_Software.ItemsSource = Control.Management.Software_agms3;
            }));

            DataGridAGMS4_Software.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate ()
            {
                DataGridAGMS4_Software.ItemsSource = null;
                DataGridAGMS4_Software.ItemsSource = Control.Management.Software_agms4;
            }));
        }


        public bool CanConnect(int selected_server, out Model.User _user)
        {
            DateTime currentTime = DateTime.Now;
            TimeSpan difference = currentTime - Management.last_refresh;
            if (difference.TotalSeconds < Management.INTERVAL_TIME_REFRESH_SECONDS)
            {
                Connection.ReadSheets();
                Management.last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = Management.last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }

            while (Management.Users_agms2.Count == 0 || Management.Users_agms3.Count == 0)
            {
                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));
                System.Threading.Thread.Sleep(OFFSET_TIME_REFRESH_MILLISECONDS);
                Connection.ReadSheets();

                Management.last_refresh = DateTime.Now;
                LastUpdate.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { LastUpdate.Text = Management.last_refresh.ToString("dd/MM/yyyy HH:mm:ss"); }));
            }
            //ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
            LoadDatagrid();

            if (selected_server == 2 || selected_server == 1)
            {
                _user = Management.Users_agms2.Where(a => a.IsLogged).FirstOrDefault();
                if (_user != null)
                    return true;
                else
                    return false;
            }
            else if (selected_server == 3)
            {
                _user = Management.Users_agms3.Where(a => a.IsLogged).FirstOrDefault();
                if (_user != null)
                    return true;
                else
                    return false;
            }
            else if (selected_server == 4)
            {
                _user = Management.Users_agms4.Where(a => a.IsLogged).FirstOrDefault();
                if (_user != null)
                    return true;
                else
                    return false;
            }

            _user = null;
            return true;
        }

        [STAThread]
        private void ProcessConnection()
        {
            Model.User _user;
            bool hasConnectedUser = CanConnect(selected_server, out _user);

            if (!hasConnectedUser || _user.Name.Equals(current_user))
            {
                //Update google sheets
                Connection.UpdateSheet(selected_server, true, false);

                string username = "sysfliu";
                string password = "Agms2023!";
                Management.ConnectAGMS(username, password, out error_taken_time_connected);
                dispatcherTimer_checkProcessAlive.Start();
            }
            else
            {
                System.Windows.MessageBox.Show(
                        "The following user is connected on AGMS" + selected_server + ": " + (_user != null ? _user.Name : "---") + "!\nPlease contact him/her to log out the server or try again later!",
                        "The Liu Lab :: Information",
                        (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                        (System.Windows.MessageBoxImage)MessageBoxIcon.Warning);
                ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = true; }));
                dispatcherTimer.Start();
                progressbar_time = 0;
            }
        }

        private bool isMaintenance()
        {
            Connection.ReadSheets();
            Model.Server? current_server = Management.Servers.Where(a => a.Index == selected_server).FirstOrDefault();
            if (current_server == null) return false;

            if (current_server.IsMaintenace == true)
                return true;
            return false;
        }

        private async void ButtonConnect_Click(object sender, RoutedEventArgs e)
        {
            dispatcherTimer.Stop();
            progressbar_time = int.MinValue;

            selected_server = -1;
            if (AGMS_tab.SelectedIndex == 0)//AGMS2
            {
                Management.IP_address = "10.10.65.31";
                selected_server = 2;
            }
            else if (AGMS_tab.SelectedIndex == 1)//AGMS3
            {
                Management.IP_address = "10.10.65.49";
                selected_server = 3;
            }
            else if (AGMS_tab.SelectedIndex == 2)//AGMS4
            {
                Management.IP_address = "10.10.65.40";
                selected_server = 4;
            }

            if(isMaintenance())
            {
                System.Windows.MessageBox.Show(
                "AGMS" + selected_server + " is under maintenance. Contact the administrator for more information.",
                "The Liu Lab :: Information",
                (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                        (System.Windows.MessageBoxImage)MessageBoxIcon.Warning);
                return;
            }

            ButtonConnect.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate () { ButtonConnect.IsEnabled = false; }));

            var wait_screen = Util.Util.CallWaitWindow("Please wait", "Checking users status...");
            MainGrid.Children.Add(wait_screen);
            Grid.SetRowSpan(wait_screen, 4);
            Grid.SetRow(wait_screen, 0);
            var rows = MainGrid.RowDefinitions;
            rows[2].Height = new GridLength(2, GridUnitType.Star);
            await Task.Run(() => ProcessConnection());

            MainGrid.Children.Remove(wait_screen);
            rows[2].Height = GridLength.Auto;
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
                else if (AGMS_tab.SelectedIndex == 2)//AGMS4
                    selected_server = 4;

                var response = System.Windows.Forms.MessageBox.Show($"Has your task been finished?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (response == System.Windows.Forms.DialogResult.Yes)
                {
                    //Update google sheets
                    Connection.UpdateSheet(selected_server, false, true);
                }
                else
                {
                    //Update google sheets
                    Connection.UpdateSheet(selected_server, false, false);
                }

                LoadDatagrid();

                progressbar_time = 0;
                dispatcherTimer.Start();

                System.Windows.MessageBox.Show(
                            "The server has been closed successfully!",
                            "The Liu Lab :: Information",
                            (System.Windows.MessageBoxButton)MessageBoxButtons.OK,
                            (System.Windows.MessageBoxImage)MessageBoxIcon.Information);
            }));

            Management.IsConnected = false;
            Management.CurrentPID = -1;
        }

        private void DataGridAGMS2_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void DataGridAGMS3_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void DataGridAGMS4_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }
    }
}
