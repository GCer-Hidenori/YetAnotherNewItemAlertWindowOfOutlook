using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using YetAnotherNewItemAlertWindowOfOutlook.Properties;
using System.Timers;
using System.Net;
using System.DirectoryServices.ActiveDirectory;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class MainViewModel
    {
        public ObservableCollection<OutlookMailItem> OutlookMailItemCollection { get; set; }
    

        public Dictionary<string, OutlookMailItem> DicOutlookMailItem = new();
        private Timer timer;
        private int timer_count = 0;
        List<TargetProcessing> list_target_processing = new();
        private Window window;
        Microsoft.Office.Interop.Outlook.Application outlook;
        private Setting setting;
        private NLog.Logger logger;

        static List<TargetProcessing> GetTargetProcessings(Setting setting,NLog.Logger logger)
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
                    var target_processing = new TargetProcessing(logger)
                    {
                        Target = target,
                        Target_folder = folder
                    };
                    list_target_processing.Add(target_processing);
                }
            }
            return list_target_processing;
        }

        public MainViewModel(Setting setting, Microsoft.Office.Interop.Outlook.Application outlook,Window window,NLog.Logger logger)
        {
            this.outlook = outlook;
            this.window = window;
            this.setting = setting;
            this.logger = logger;
            OutlookMailItemCollection = new ObservableCollection<OutlookMailItem>();
            list_target_processing = MainViewModel.GetTargetProcessings(setting,logger);

            RefreshOutlookMailItem();
            StartTimer();
            
        }
        public void StartTimer()
        {
 
            int timer_interval_millisec = setting.TimerIntervalSec * 1000;
            timer = new Timer(timer_interval_millisec); 
            timer.Elapsed += (sender, e) =>
            {
                RefreshOutlookMailItem();
            };
            timer.Start();
        }
        object lockObj = new();

        public void RefreshOutlookMailItem(bool forceRefresh=false)
        {
            lock (lockObj)
            {
                System.Diagnostics.Debug.WriteLine("RefreshOutlookMailItem");
                logger.Info("start RefreshOutlookMailItem");
                bool activateWindow = false;

                List<string> list_entryid_of_target_processing = new();
                List<string> list_entryid_of_outlookmailitemcollection = new();
                foreach (var target_processing in list_target_processing)
                {
                    if (forceRefresh || ( target_processing.Target.IntervalMin > 0 && timer_count % target_processing.Target.IntervalMin == 0))
                    {
                        var result = target_processing.RefreshOutlookMailItem();
                        if(!activateWindow && result.ActivateWindow) activateWindow = true;
                        foreach (string entryID in result.List_new_entry_id)
                        {
                            foreach (ActionCreateFile actionCreateFile in target_processing.Target.ActionCreateFiles)
                            {
                                MailItem mailItem = OutlookUtil.GetMail(entryID,outlook);
                                actionCreateFile.CreateFile(mailItem);
                            }
                        }
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
                if(activateWindow)
                {
                    logger.Info("activate window.");
                    window.Dispatcher.Invoke(() => {
                        window.Activate();
                        window.WindowState = WindowState.Normal;
                    });
                }
            }

        }
        public void PauseTimer()
        {
            timer.Stop();
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
