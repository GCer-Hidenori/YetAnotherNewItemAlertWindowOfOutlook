using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Timers;
using System.Windows;
using System.Windows.Data;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class MainViewModel
    {
        public ObservableCollection<OutlookMailItem> OutlookMailItemCollection { get; set; }

        public Dictionary<MailID, OutlookMailItem> DicOutlookMailItem = new();
        private Timer? timer;
        private int timer_count = 0;
        List<TargetProcessing> list_target_processing = new();
        private Window window;
        private Setting setting;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        static List<TargetProcessing> GetTargetProcessings(Setting setting)
        {
            List<TargetProcessing> list_target_processing = new();
            foreach (Target target in setting.Targets)
            {
                var target_processing = new TargetProcessing(target);
                list_target_processing.Add(target_processing);
            }
            return list_target_processing;
        }

        public MainViewModel(Setting setting, Window window, IgnoreFileList ignoreFileList)
        {
            this.window = window;
            this.setting = setting;
            OutlookMailItemCollection = new ObservableCollection<OutlookMailItem>();
            BindingOperations.EnableCollectionSynchronization(OutlookMailItemCollection, new object());
            list_target_processing = MainViewModel.GetTargetProcessings(setting);

            RefreshOutlookMailItem(ignoreFileList);

            SetTimer(ignoreFileList);
            StartTimer();
        }
        private void SetTimer(IgnoreFileList ignoreFileList)
        {
            int timer_interval_millisec = setting.TimerIntervalSec * 1000;
            timer = new Timer(timer_interval_millisec);
            timer.Elapsed += (sender, e) =>
            {
                RefreshOutlookMailItem(ignoreFileList);
            };
            timer.AutoReset = true;
        }

        public void StartTimer()
        {
            Logger.Info("start timer.");
            timer?.Start();
        }
        object lockObj = new();

        public void RefreshOutlookMailItem(IgnoreFileList ignoreFileList, bool forceRefresh = false)
        {
            try
            {
                var outlook = new Microsoft.Office.Interop.Outlook.Application();
                lock (lockObj)
                {
                    System.Diagnostics.Debug.WriteLine("RefreshOutlookMailItem start");

                    List<MailID> list_entryid_of_target_processing = new();
                    List<MailID> list_entryid_of_outlookMailItemcollection = new();
                    foreach (var target_processing in list_target_processing)
                    {
                        if (forceRefresh || (target_processing.Target?.TimersToCheckMail > 0 && timer_count % target_processing.Target.TimersToCheckMail == 0))
                        {
                            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToString()} start RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                            Logger.Info($"start RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                            var result = target_processing.RefreshOutlookMailItem(ignoreFileList, window);
                            Logger.Info($"end RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                        }

                        foreach (MailID mailID in target_processing.List_OutlookMailID)
                        {
                            list_entryid_of_target_processing.Add(new MailID() { StoreID = mailID.StoreID, EntryID = mailID.EntryID });
                        }

                    }

                    foreach (OutlookMailItem outlookMailItem in OutlookMailItemCollection)
                    {
                        list_entryid_of_outlookMailItemcollection.Add(new MailID() { StoreID = outlookMailItem.StoreID, EntryID = outlookMailItem.EntryID });
                    }


                    // duplicate / for update
                    List<MailID> list_duplication_entryid = list_entryid_of_target_processing.Intersect(list_entryid_of_outlookMailItemcollection).ToList();
                    // only target_processing / add to OutlookMailItemCollection
                    List<MailID> list_new_entry_id = list_entryid_of_target_processing.Except(list_entryid_of_outlookMailItemcollection).ToList();
                    // only OutlookMailItemCollection / delete from OutlookMailItemCollection
                    List<MailID> list_deleted_entry_id = list_entryid_of_outlookMailItemcollection.Except(list_entryid_of_target_processing).ToList();

                    Logger.Info($"new item  count  {list_new_entry_id.Count}");

                    foreach (MailID mailID in list_duplication_entryid)
                    {
                        try
                        {
                            OutlookMailItem.Reload(DicOutlookMailItem[mailID], outlook);
                        }
                        catch (System.Runtime.InteropServices.COMException e)
                        {
                            Logger.Warn(e);
                        }
                    }
                    foreach (MailID mailID in list_new_entry_id)
                    {
                        var outlookMailItem = OutlookMailItem.CreateNew(mailID.StoreID, mailID.EntryID, outlook);
                        DicOutlookMailItem.Add(mailID, outlookMailItem);
                        OutlookMailItemCollection.Add(outlookMailItem);
                    }
                    foreach (MailID mailID in list_deleted_entry_id)
                    {
                        var outlookMailItem = DicOutlookMailItem[mailID];
                        DicOutlookMailItem.Remove(mailID);
                        if (!OutlookMailItemCollection.Remove(outlookMailItem))
                        {
                            throw new ArgumentException();
                        }
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    timer_count++;
                }
            }
            catch (System.Exception e)
            {
                Logger.Error(e);
            }

        }
        public void StopTimer()
        {
            Logger.Info("stop timer.");
            timer?.Stop();
        }
        public void HideMail(string entryID, string storeID)
        {
            MailID mailID = new() { StoreID = storeID, EntryID = entryID };
            foreach (var target_processing in list_target_processing)
            {
                if (target_processing.List_OutlookMailID.Contains(mailID))
                {
                    target_processing.List_OutlookMailID.Remove(mailID);
                }
            }
            var outlookMailItem = DicOutlookMailItem[mailID];
            DicOutlookMailItem.Remove(mailID);
            OutlookMailItemCollection.Remove(outlookMailItem);
        }
    }
}
