using System;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Outlook;
//using System.Windows.Forms;
using System.Windows.Interop;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MainViewModel context;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private bool ready = false;
        string settingFilePath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "setting.xml");


        public MainWindow()
        {
            InitializeComponent();
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            try
            {
                Setting setting = Setting.Load(Logger);
                context = new MainViewModel(setting, outlook, this,Logger);
                this.DataContext = context;
                ready = true;
            }
            catch (YError e)
            {
                switch (e.ErrorType)
                {
                    case ErrorType.SettingFileNotFound:
                        Logger.Error("Setting file not found.");
                        Logger.Error(e.Message);
                        Logger.Info("sample setting file:\n" + Util.ReadSettingSampleXmlString());
                        MessageBox.Show("Setting file not found.see log file for sample of setting.xml");
                        MessageBoxResult res = MessageBox.Show("Do you want to create a configuration file from your Outlook folder structure?", "Confirmation", MessageBoxButton.YesNo);
                        if (res == MessageBoxResult.Yes)
                        {
                            try
                            {
                                Util.CreateSettingFile(outlook,settingFilePath,Logger);
                                Logger.Info("Setting file created.");
                                MessageBox.Show("Setting file created.\nExit the application. Restart the application manually.");
                                this.Close();
                            }
                            catch (System.Exception ex)
                            {
                                Logger.Error("Setting file creation error.");
                                Logger.Error(ex.Message);
                                MessageBox.Show("Setting file creation error.");
                            }
                        }
                        else
                        {
                            Logger.Info("User canceled.");
                        }
                        break;
                    case ErrorType.SettingFileLoadError:
                        Logger.Error("Setting file load error.");
                        Logger.Error(e.Message);
                        Logger.Info("sample setting file:\n" + Util.ReadSettingSampleXmlString());
                        MessageBox.Show("Setting file load error.see log file for sample of setting.xml");
                        break;
                    default:
                        MessageBox.Show(e.Message);
                        Logger.Error(e.Message);
                        break;
                }
                try
                {
                    if (context != null)
                    {
                        context.PauseTimer();
                    }
                }
                catch (System.Exception)
                {

                }
                
          
                //this.Close();
            }

        }
        /*
        public static void Dialog(string message)
        {
            MessageBoxResult res = MessageBox.Show(message, "Confirmation", MessageBoxButton.OK);
        }
        */
        private void DataGridRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            var mailItem = ns.GetItemFromID(((OutlookMailItem)((DataGridRow)sender).Item).EntryID);
            if (mailItem != null)
            {
                mailItem.Display();
                mailItem.GetInspector.Display(false);
            }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            /*
            if (this.WindowState == WindowState.Minimized)
            {
                this.Hide();
            }
            */
        }

        private void RefreshNow_Click(object sender, RoutedEventArgs e)
        {
            if(ready)context.RefreshOutlookMailItem(true);
        }
        private void PauseTimer_Click(object sender, RoutedEventArgs e)
        {
            if (ready) context.PauseTimer();
        }
        private void StartTimer_Click(object sender, RoutedEventArgs e)
        {
            if (ready) context.StartTimer();
        }

        private void DataGridRow_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Delete)
            {
                System.Diagnostics.Debug.WriteLine("del key ");
                OutlookMailItem outlookMailItem = (OutlookMailItem)((DataGridRow)sender).Item;
                string entryID = outlookMailItem.EntryID;
                IgnoreFile.Add(entryID,outlookMailItem,Logger);
                context.HideMail(entryID);
            }
        }

        private void ListFolders_Click(object sender,RoutedEventArgs e)
        {
            OutlookUtil.ListAllFolders(Logger);
        }
        private void OpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            string logDir = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "log");
            if(!Directory.Exists(logDir))Directory.CreateDirectory(logDir);
            var psi = new System.Diagnostics.ProcessStartInfo() { FileName = logDir, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }
        private void OpenSettingFile_Click(object sender, RoutedEventArgs e)
        {
            var psi = new System.Diagnostics.ProcessStartInfo() { FileName = settingFilePath, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }
    }
}
