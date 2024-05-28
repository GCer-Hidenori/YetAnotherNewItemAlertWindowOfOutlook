using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public struct MailID
    {
        public string StoreID;
        public string EntryID;
    }
    internal class TargetProcessing
    {
        private Target target;
        private List<MailID> list_outlookmaili_entryID = new();
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public TargetProcessing(Target target)
        {
            this.target = target;
        }

        public Target Target { get => target; set => target = value; }
        //public MAPIFolder? Target_folder { get => target_folder; set => target_folder = value; }

        public MAPIFolder GetTargetFolder()
        {
            if (target.TargetFolderType == Target.FolderType.SearchFolder)
            {
                return OutlookUtil.GetSearchFolder(target.Path);
            }
            else
            {
                return OutlookUtil.GetNormalFolder(target.Path);
            }
        }

        public List<MailID> List_OutlookMailEntryID { get => list_outlookmaili_entryID; set => list_outlookmaili_entryID = value; }


        public ResultOfTargetProcessing RefreshOutlookMailItem(IgnoreFileList ignoreFileList)
        {
            var result = new ResultOfTargetProcessing();
            List<MailID> original_list_outlookmaili_entryID = new(list_outlookmaili_entryID);
            list_outlookmaili_entryID.Clear();
            var folder = GetTargetFolder();
            foreach (object item in folder.Items)
            {
                if (item is MailItem mailItem)
                {
                    if (ignoreFileList.Exists(mailItem.Parent.StoreID,mailItem.EntryID))
                    {
                        continue;
                    }
                    else
                    {
                        if (!target.Filtering(mailItem))
                        {
                            continue;
                        }
                    }
                    MailID mailID = new() { StoreID = mailItem.Parent.StoreID, EntryID = mailItem.EntryID };
                    List_OutlookMailEntryID.Add(mailID);
                    if (result.ActivateWindow == false && target.ActivateWindow && !original_list_outlookmaili_entryID.Contains(mailID))
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
