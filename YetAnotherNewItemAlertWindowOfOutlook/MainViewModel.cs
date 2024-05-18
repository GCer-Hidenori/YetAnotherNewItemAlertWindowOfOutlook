using Microsoft.Office.Interop.Outlook;
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

        public Dictionary<string, OutlookMailItem> DicOutlookMailItem = new();
        private Timer? timer;
        private int timer_count = 0;
        List<TargetProcessing> list_target_processing = new();
        private Window window;
        //Microsoft.Office.Interop.Outlook.Application outlook;
        private Setting setting;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        static List<TargetProcessing> GetTargetProcessings(Setting setting)
        {
            List<TargetProcessing> list_target_processing = new();
            foreach (Target target in setting.Targets)
            {
                MAPIFolder? folder = null;
                if (target.TargetFolderType == Target.FolderType.SearchFolder)
                {
                    if (target.Path != null)
                    {
                        folder = OutlookUtil.GetSearchFolder(target.Path);
                    }
                }
                else
                {
                    //TBW
                    if (target.Path != null)
                    {
                        folder = OutlookUtil.GetNormalFolder(target.Path);
                    }
                }
                if (folder != null)
                {
                    var target_processing = new TargetProcessing()
                    {
                        Target = target,
                        Target_folder = folder
                    };
                    list_target_processing.Add(target_processing);
                }
            }
            return list_target_processing;
        }

        public MainViewModel(Setting setting,Window window)
        {
            //this.outlook = outlook;
            this.window = window;
            this.setting = setting;
            OutlookMailItemCollection = new ObservableCollection<OutlookMailItem>();
            BindingOperations.EnableCollectionSynchronization(OutlookMailItemCollection, new object());
            list_target_processing = MainViewModel.GetTargetProcessings(setting);

            RefreshOutlookMailItem();

            SetTimer();
            StartTimer();

        }
        private void SetTimer()
        {
            int timer_interval_millisec = setting.TimerIntervalSec * 1000;
            timer = new Timer(timer_interval_millisec);
            timer.Elapsed += (sender, e) =>
            {
                RefreshOutlookMailItem();
            };
            timer.AutoReset = true;

        }


        public void StartTimer()
        {
            Logger.Info("start timer.");
            timer?.Start();
        }
        object lockObj = new();

        public void RefreshOutlookMailItem(bool forceRefresh = false)
        {
            try
            {
                var outlook = new Microsoft.Office.Interop.Outlook.Application();
                lock (lockObj)
                {
                    System.Diagnostics.Debug.WriteLine("RefreshOutlookMailItem start");

                    bool activateWindow = false;

                    List<string> list_entryid_of_target_processing = new();
                    List<string> list_entryid_of_outlookmailitemcollection = new();
                    foreach (var target_processing in list_target_processing)
                    {
                        if (forceRefresh || (target_processing.Target?.TimersToCheckMail > 0 && timer_count % target_processing.Target.TimersToCheckMail == 0))
                        {
                            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToString()} start RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                            Logger.Info($"start RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                            var result = target_processing.RefreshOutlookMailItem();

                            if (!activateWindow && result.ActivateWindow) activateWindow = true;
                            foreach (string entryID in result.List_new_entry_id)
                            {
                                if (target_processing.Target != null)
                                {
                                    foreach (Action actionCreateFile in target_processing.Target.Actions.Where(a => a.ActionType == ActionType.CreateFile))
                                    {

                                        MailItem mailItem = OutlookUtil.GetMail(entryID, outlook);
                                        actionCreateFile.CreateFile(mailItem);
                                    }
                                }
                            }
                            Logger.Info($"end RefreshOutlookMailItem.Folder:{target_processing.Target?.Path}");
                        }

                        foreach (string mail_entryID in target_processing.List_OutlookMailEntryID)
                        {
                            list_entryid_of_target_processing.Add(mail_entryID);
                        }

                    }

                    foreach (OutlookMailItem outlookmailitem in OutlookMailItemCollection)
                    {
                        list_entryid_of_outlookmailitemcollection.Add(outlookmailitem.EntryID);
                    }


                    // duplicate / for update
                    List<string> list_duplication_entryid = list_entryid_of_target_processing.Intersect(list_entryid_of_outlookmailitemcollection).ToList();
                    // only target_processing / add to OutlookMailItemCollection
                    List<string> list_new_entry_id = list_entryid_of_target_processing.Except(list_entryid_of_outlookmailitemcollection).ToList();
                    // only OutlookMailItemCollection / delete from OutlookMailItemCollection
                    List<string> list_deleted_entry_id = list_entryid_of_outlookmailitemcollection.Except(list_entryid_of_target_processing).ToList();

                    Logger.Info($"new item  count  {list_new_entry_id.Count}");

                    foreach (string entryID in list_duplication_entryid)
                    {
                        OutlookMailItem.Reload(DicOutlookMailItem[entryID], outlook);
                    }
                    foreach (string entryID in list_new_entry_id)
                    {
                        var outlookmailitem = OutlookMailItem.CreateNew(entryID, outlook);
                        DicOutlookMailItem.Add(entryID, outlookmailitem);
                        OutlookMailItemCollection.Add(outlookmailitem);
                    }
                    foreach (string entryID in list_deleted_entry_id)
                    {
                        var outlookmailitem = DicOutlookMailItem[entryID];
                        DicOutlookMailItem.Remove(entryID);
                        if (!OutlookMailItemCollection.Remove(outlookmailitem))
                        {
                            throw new ArgumentException();
                        }
                    }
                    timer_count++;
                    if (activateWindow)
                    {
                        Logger.Info("activate window.");
                        window.Dispatcher.Invoke(() =>
                        {
                            window.Activate();
                            window.WindowState = WindowState.Normal;
                        });
                    }
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
        public void HideMail(string entryID)
        {
            foreach (var target_processing in list_target_processing)
            {
                if (target_processing.List_OutlookMailEntryID.Contains(entryID))
                {
                    target_processing.List_OutlookMailEntryID.Remove(entryID);
                }
            }
            var outlookmailitem = DicOutlookMailItem[entryID];
            DicOutlookMailItem.Remove(entryID);
            OutlookMailItemCollection.Remove(outlookmailitem);
        }
    }
}
