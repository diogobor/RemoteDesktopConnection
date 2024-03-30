using RemoteDesktopConnection.Control;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace RemoteDesktopConnection.Viewer
{
    /// <summary>
    /// Interaction logic for IdleWindow.xaml
    /// </summary>
    public partial class IdleWindow : Window
    {
        private System.Windows.Threading.DispatcherTimer dispatcherProgressBar = new System.Windows.Threading.DispatcherTimer();
        private double progressbar_time = 0;

        public bool CloseSoftware { get;set; }
        public IdleWindow()
        {
            InitializeComponent();

            dispatcherProgressBar.Tick += new EventHandler(dispatcherProgressBar_Tick);
            dispatcherProgressBar.Interval = new TimeSpan(0, 0, 1);
            dispatcherProgressBar.Start();
        }

        private void dispatcherProgressBar_Tick(object sender, EventArgs e)
        {
            if (progressbar_time < Management.INTERVAL_TIME_REFRESH_SECONDS)
            {
                progressbar_time++;
                ProgressBarRefresh.Value = progressbar_time * (100.0 / Management.INTERVAL_TIME_REFRESH_SECONDS);
            }
            else
            {
                //Close Software
                CloseSoftware = true;
                this.Close();
            }
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            CloseSoftware = false;
            this.Close();
        }
    }
}
