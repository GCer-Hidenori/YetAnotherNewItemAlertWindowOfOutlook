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
        private static string sample_setting_xml = @"<Setting>
    <Targets>
        <Target>
            <TargetFolderType>NormalFolder</TargetFolderType>
            <IntervalMin>1</IntervalMin>
            <Path>\\yanaw@example.com\ImportantMails</Path>
        </Target>
        <Target>
            <TargetFolderType>SearchFolder</TargetFolderType>
            <IntervalMin>1</IntervalMin>
            <Path>\\yanaw@example.com\search folder\search folder01</Path>
		    <Filter>
			    <And>
				    <SenderEmailAddress>sender111@example.com</SenderEmailAddress>
			    </And>
		    </Filter>
            <Actions>
                <Activate_Window>true</Activate_Window>
                <Create_File fileName=""c:\work\inbox\${entryID}.md"">
                    <body>---
Subject: ${Subject}
ReceivedTime: ${ReceivedTime}
From: ${SenderName}&lt;${SenderEmailAddress}&gt;
---
# Body
${Body}</body>
                </Create_File>
            </Actions>
        </Target>
    </Targets>
</Setting>"; 
        MainViewModel context;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public MainWindow()
        {
            InitializeComponent();
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            try
            {
                Setting setting = Setting.Load(Logger);
                context = new MainViewModel(setting, outlook, this,Logger);
                this.DataContext = context;
            }
            catch (YError e)
            {
                switch (e.ErrorType)
                {
                    case ErrorType.SettingFileNotFound:
                        MessageBox.Show("Setting file not found.see log file for sample of setting.xml");
                        Logger.Error("Setting file not found.");
                        Logger.Info("sample setting file:\n" + sample_setting_xml);
                        break;
                    case ErrorType.SettingFileLoadError:
                        MessageBox.Show("Setting file not found.see log file for sample of setting.xml");
                        Logger.Error("Setting file not found.");
                        Logger.Info("sample setting file:\n" + sample_setting_xml);
                        break;
                    default:
                        MessageBox.Show(e.Message);
                        Logger.Error(e.Message);
                        break;
                }
          
                this.Close();
            }

        }
        public static void Dialog(string message)
        {
            MessageBoxResult res = MessageBox.Show(message, "Confirmation", MessageBoxButton.OK);
        }
        private void DataGridRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            ns.GetItemFromID(((OutlookMailItem)((DataGridRow)sender).Item).EntryID).Display();
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
            context.RefreshOutlookMailItem(true);
        }
        private void PauseTimer_Click(object sender, RoutedEventArgs e)
        {
            context.PauseTimer();
        }
        private void StartTimer_Click(object sender, RoutedEventArgs e)
        {
            context.StartTimer();
        }

        private void DataGridRow_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Delete)
            {
                System.Diagnostics.Debug.WriteLine("del key ");
                string entryID = ((OutlookMailItem)((DataGridRow)sender).Item).EntryID;
                IgnoreFile.Add(entryID);
                context.HideMail(entryID);
            }
        }
    }
}
