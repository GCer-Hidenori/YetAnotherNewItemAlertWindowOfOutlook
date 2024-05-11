using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

using System.Printing;
using NLog;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class TargetProcessing
    {
        private Target target;
        private MAPIFolder target_folder;
        private List<string> list_outlookmaili_entryID = new();
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public Target Target { get => target; set => target = value; }
        public MAPIFolder Target_folder { get => target_folder; set => target_folder = value; }

        public List<string> List_OutlookMailEntryID {get => list_outlookmaili_entryID; set => list_outlookmaili_entryID = value; }


        public ResultOfTargetProcessing RefreshOutlookMailItem()
        {
            var result = new ResultOfTargetProcessing();
            List<string> original_list_outlookmaili_entryID = new(list_outlookmaili_entryID);
            list_outlookmaili_entryID.Clear();
            foreach (object item in target_folder.Items)
            {
                if (item is MailItem mailItem)
                {
                    if(IgnoreFile.Exists(mailItem.EntryID))
                    {
                         continue; 
                    }else 
                    {
                        if(!target.Filtering(mailItem))
                        {
                            continue;
                        }
                    }
                    
                    List_OutlookMailEntryID.Add(mailItem.EntryID);
                    if(result.ActivateWindow == false && target.ActivateWindow && !original_list_outlookmaili_entryID.Contains(mailItem.EntryID))
                    {
                        result.ActivateWindow = true;
                    }
                }
            }
            result.List_new_entry_id = list_outlookmaili_entryID.Except(original_list_outlookmaili_entryID).ToList();
            return result;
        }   
    }   
}
