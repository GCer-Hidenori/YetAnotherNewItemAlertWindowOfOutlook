using Microsoft.Office.Interop.Outlook;
//using System.Windows.Forms;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
namespace YetAnotherNewItemAlertWindowOfOutlook
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MainViewModel? context;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private bool ready = false;
        Setting setting;
        DataGrid datagrid;
        IgnoreFileList ignoreFileList = IgnoreFileList.Init();

        private void SortColumn()
        {
            foreach (var column in setting.Columns)
            {
                var targetColumn = datagrid?.Columns.FirstOrDefault(c => c.Header.ToString() == column.Name);
                if (targetColumn != null)
                {
                    targetColumn.DisplayIndex = setting.Columns.IndexOf(column);
                    if (column.Width != null)
                    {
                        targetColumn.Width = new DataGridLength((double)column.Width);
                    }

                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            datagrid = (DataGrid)this.FindName("OutlookMailItemDataGrid");

            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            try
            {
                setting = Setting.Init();
                context = new MainViewModel(setting, this, ignoreFileList);
                this.DataContext = context;
                SortColumn();



                ready = true;
            }
            catch (YError e)
            {
                switch (e.ErrorType)
                {
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
                        context.StopTimer();
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
            OpenMailItem();
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
            if (ready) context?.RefreshOutlookMailItem(ignoreFileList, true);
        }
        private void StopTimer_Click(object sender, RoutedEventArgs e)
        {
            if (ready) context?.StopTimer();
        }
        private void StartTimer_Click(object sender, RoutedEventArgs e)
        {
            if (ready) context?.StartTimer();
        }

        private void HideItemByEvent()
        {
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                string entryID = outlookMailItem.EntryID;
                ignoreFileList.Add(outlookMailItem.StoreID, entryID);
                //IgnoreFile.Add(outlookMailItem.StoreID, entryID, outlookMailItem, Logger);
                context?.HideMail(entryID, outlookMailItem.StoreID);
            }

            //OutlookMailItem outlookMailItem = (OutlookMailItem)((DataGridRow)sender).Item;
            //string entryID = outlookMailItem.EntryID;
            //IgnoreFile.Add(entryID, outlookMailItem, Logger);
            //context?.HideMail(entryID);
        }
        private void DeleteItemByEvent()
        {
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                string entryID = outlookMailItem.EntryID;
                context?.HideMail(entryID, outlookMailItem.StoreID);
            }

        }
        private void DataGridRow_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Insert:
                    HideItemByEvent();
                    break;
                case Key.Delete:
                    if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                    {
                        DeleteFromOutlook();    //delete from outlook
                    }
                    else
                    {
                        DeleteItemByEvent();    //only delete from this app
                    }
                    break;

                default:
                    break;
            }
        }

        private void ListFolders_Click(object sender, RoutedEventArgs e)
        {
            OutlookUtil.ListAllFolders(Logger);
        }
        private void OpenIgnoreListFile_Click(object sender, RoutedEventArgs e)
        {
            string ignoreListFilePath = IgnoreFileList.ignore_file_list_path;
            if (File.Exists(ignoreListFilePath))
            {
                var psi = new System.Diagnostics.ProcessStartInfo() { FileName = ignoreListFilePath, UseShellExecute = true };
                System.Diagnostics.Process.Start(psi);
            }
            else
            {
                MessageBox.Show("There is no ignore list file yet.");
            }
        }
        private void OpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            string logDir = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "log");
            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
            var psi = new System.Diagnostics.ProcessStartInfo() { FileName = logDir, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }
        private void ClearIgnoreList_Click(object sender, RoutedEventArgs e)
        {
            ignoreFileList = new IgnoreFileList();
        }
        private void DeleteUnwantedIgnoreList_Click(object sender, RoutedEventArgs e)
        {
            ignoreFileList.DeleteUnwantedIgnoreList();
        }
        private void OpenSettingFile_Click(object sender, RoutedEventArgs e)
        {
            var psi = new System.Diagnostics.ProcessStartInfo() { FileName = Setting.fileName, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }
        public bool Contains(object de)
        {
            OutlookMailItem outlookMailItem = (OutlookMailItem)de;
            var textbox = (TextBox)this.FindName("SearchTextBox");
            return textbox.Text.Split(' ').All(word => outlookMailItem.SearchIndex.Contains(Strings.StrConv(word, VbStrConv.Wide) ?? "", StringComparison.CurrentCultureIgnoreCase));
        }

        private void Search()
        {
            var view = CollectionViewSource.GetDefaultView(context?.OutlookMailItemCollection);
            view.Filter = Contains;
        }
        private void SearchCancel()
        {
            var view = CollectionViewSource.GetDefaultView(context?.OutlookMailItemCollection);
            var textbox = (TextBox)this.FindName("SearchTextBox");
            textbox.Text = "";
            view.Filter = null;
        }

        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Escape:
                    SearchCancel();
                    break;
                default:
                    break;
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ready)
            {
                var textbox = (TextBox)this.FindName("SearchTextBox");
                if (textbox.Text == "")
                {
                    SearchCancel();
                }
                else
                {
                    Search();
                }
            }

        }

        private void SearchTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (ready)
            {
                var textbox = (TextBox)this.FindName("SearchTextBox");
                if (textbox.Foreground != new SolidColorBrush(Colors.Black))
                {
                    textbox.Text = "";
                    textbox.Foreground = new SolidColorBrush(Colors.Black);
                    textbox.GotFocus -= SearchTextBox_GotFocus;
                }
            }
        }

        private void OpenMailItem1()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                try
                {

                    MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                    if (mailItem != null)
                    {
                        mailItem.Display();
                        System.Diagnostics.Process[] ps =
    System.Diagnostics.Process.GetProcessesByName("outlook");
                        if (0 < ps.Length)
                        {
                            Microsoft.VisualBasic.Interaction.AppActivate(ps[0].Id);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException e2)
                {
                    MessageBox.Show("Can't open mail.");
                    Logger.Warn(e2);
                }
            }
        }
        private void OpenMailItem2()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                try
                {

                    MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                    if (mailItem != null)
                    {
                        mailItem.GetInspector.Display(false);
                        System.Diagnostics.Process[] ps =
    System.Diagnostics.Process.GetProcessesByName("outlook");
                        if (0 < ps.Length)
                        {
                            Microsoft.VisualBasic.Interaction.AppActivate(ps[0].Id);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException e2)
                {
                    MessageBox.Show("Can't open mail.");
                    Logger.Warn(e2);
                }
            }
        }
        private void OpenMailItem3()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                try
                {

                    MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                    if (mailItem != null)
                    {
                        mailItem.GetInspector.Activate();
                        System.Diagnostics.Process[] ps =
    System.Diagnostics.Process.GetProcessesByName("outlook");
                        if (0 < ps.Length)
                        {
                            Microsoft.VisualBasic.Interaction.AppActivate(ps[0].Id);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException e2)
                {
                    MessageBox.Show("Can't open mail.");
                    Logger.Warn(e2);
                }
            }
        }
        private void OpenMailItem()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                try
                {

                    MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                    if (mailItem != null)
                    {
                        mailItem.Display();
                        mailItem.GetInspector.Display(false);
                        mailItem.GetInspector.Activate();
                        
                        System.Diagnostics.Process[] ps =
                            System.Diagnostics.Process.GetProcessesByName("outlook");
                        if (0 < ps.Length)
                        {
                            Microsoft.VisualBasic.Interaction.AppActivate(ps[0].Id);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException e2)
                {
                    MessageBox.Show("Can't open mail.");
                    Logger.Warn(e2);
                }
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            setting.Columns.Clear();
            foreach (var column in datagrid.Columns.OrderBy(c => c.DisplayIndex))
            {
                setting.Columns.Add(new Column() { Name = column.Header.ToString() ?? "", Width = column.ActualWidth });
            }
            setting.Save();
            ignoreFileList.Save();
        }

        private void OpenMenuItem_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            OpenMailItem();
        }
        private void OpenMenuItem1_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            OpenMailItem1();
        }
        private void OpenMenuItem2_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            OpenMailItem2();
        }
        private void OpenMenuItem3_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            OpenMailItem3();
        }
        private void HideMenuItem_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            HideItemByEvent();
        }
        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            DeleteItemByEvent();
        }
        private void DeleteFromOutlook()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");

            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                MailItem mailItem;
                try
                {
                    mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                }
                catch (System.Runtime.InteropServices.COMException e2)
                {
                    MessageBox.Show("Can't open mail.");
                    Logger.Warn(e2);
                    return;
                }
                MessageBoxResult res = MessageBox.Show($"Would you like to delete this mail from Outlook?", "Confirmation", MessageBoxButton.YesNoCancel);
                switch (res)
                {
                    case MessageBoxResult.Yes:
                        try
                        {
                            context?.HideMail(mailItem.EntryID, outlookMailItem.StoreID);
                            mailItem.Delete();
                        }
                        catch (System.Runtime.InteropServices.COMException e3)
                        {
                            MessageBox.Show("Can't delete mail.");
                            Logger.Warn(e3);
                        }
                        break;
                    case MessageBoxResult.Cancel:
                        MessageBox.Show("Canceled.");
                        return;
                    default:
                        break;
                }
            }
        }
        private void DeleteFromOutlookMenuItem_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            DeleteFromOutlook();
        }
        private void InspectMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            var outlookMailItem = (OutlookMailItem?)datagrid?.SelectedItem;
            if (outlookMailItem != null)
            {
                MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                e.Handled = true;
                string recipientNames = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(new Func<Recipient, string>(recipient => recipient.Name)));
                string recipientAddresses = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(new Func<Recipient, string>(recipient => recipient.Address)));

                string message = $@"Subject:{mailItem.Subject}
To:{mailItem.To}
Cc:{mailItem.CC}
SenderName:{mailItem.SenderName}
SenderAddress:{mailItem.SenderEmailAddress}
RecipientNames:{recipientNames}
RecipientAddresses:{recipientAddresses}
EntryID:{mailItem.EntryID}
ConversationID:{mailItem.ConversationID}
                ";
                MessageBox.Show(message, "Inspect", MessageBoxButton.OK);
            }
            else
            {
                MessageBox.Show("Can't open mail.");
            }
        }


        private void MoveToSameFolderSameThres_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                MailItem mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
                MAPIFolder? sameThreadMailFolder = OutlookUtil.GetSameThreadMailFolder(mailItem);
                if (sameThreadMailFolder != null)
                {
                    MessageBoxResult res = MessageBox.Show($"Would you like to move this mail to here?\n{sameThreadMailFolder.FullFolderPath}", "Confirmation", MessageBoxButton.YesNoCancel);
                    switch (res)
                    {
                        case MessageBoxResult.Yes:
                            mailItem.Move(sameThreadMailFolder);
                            context?.HideMail(outlookMailItem.EntryID, outlookMailItem.StoreID);
                            break;
                        case MessageBoxResult.Cancel:
                            MessageBox.Show("Canceled.");
                            return;
                        default:
                            break;
                    }

                }
                else
                {
                    MessageBox.Show("No same thread mail in other folder.");
                }
            }
        }
   
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.F5:
                    if (ready) context?.RefreshOutlookMailItem(ignoreFileList, true);
                    break;
                default:
                    break;
            }
        }




        /*
        private void OutlookMailItemDataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            
            if(e.Column.SortDirection == System.ComponentModel.ListSortDirection.Ascending)
            {
                e.Column.SortDirection = System.ComponentModel.ListSortDirection.Descending;
            }
            else
            {
                e.Column.SortDirection = System.ComponentModel.ListSortDirection.Ascending;
            }


            var view = CollectionViewSource.GetDefaultView(context.OutlookMailItemCollection);
            view.SortDescriptions.Clear();
            switch (e.Column.SortMemberPath)
            {
                case "ReceivedTime":
                    view.SortDescriptions.Add(new System.ComponentModel.SortDescription("ReceivedTime", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                case "Subject":
                    
                    //view.SortDescriptions.Add(new System.ComponentModel.SortDescription("Subject", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                case "SenderName":
                    view.SortDescriptions.Add(new System.ComponentModel.SortDescription("SenderName", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                case "To":
                    view.SortDescriptions.Add(new System.ComponentModel.SortDescription("To", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                case "Cc":
                    view.SortDescriptions.Add(new System.ComponentModel.SortDescription("Cc", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                case "EntryID":
                    view.SortDescriptions.Add(new System.ComponentModel.SortDescription("EntryID", e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
                    break;
                default:
                    break;

            }
            view.SortDescriptions.Add(new System.ComponentModel.SortDescription(e.Column.SortMemberPath, e.Column.SortDirection ?? System.ComponentModel.ListSortDirection.Ascending));
        }
        */
    }
}
