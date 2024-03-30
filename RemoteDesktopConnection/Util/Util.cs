using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using RemoteDesktopConnection.Viewer;
using System.Reflection;

namespace RemoteDesktopConnection.Util
{
    public static class Util
    {
        public static UCWaitScreen CallWaitWindow(string primaryMsg, string secondaryMg)
        {
            UCWaitScreen waitScreen = new UCWaitScreen(primaryMsg, secondaryMg);
            Grid.SetRow(waitScreen, 0);
            Grid.SetRowSpan(waitScreen, 2);
            waitScreen.Margin = new Thickness(0, 0, 0, 0);
            return waitScreen;
        }
   
        public static string GetVersion()
        {
            Version? version = null;
            try
            {
                version = Assembly.GetExecutingAssembly()?.GetName()?.Version;
            }
            catch (Exception e)
            {
                //Unable to retrieve version number
                Console.WriteLine("", e);
                return "";
            }
            return version.Major + "." + version.Minor;
        }
    }
}
