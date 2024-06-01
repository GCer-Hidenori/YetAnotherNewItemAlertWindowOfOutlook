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
        private List<MailID> list_outlookmail_mail_id = new();
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

        public List<MailID> List_OutlookMailID { get => list_outlookmail_mail_id; set => list_outlookmail_mail_id = value; }


        public ResultOfTargetProcessing RefreshOutlookMailItem(IgnoreFileList ignoreFileList)
        {
            var result = new ResultOfTargetProcessing();
            List<MailID> original_list_outlookmail_mail_id = new(list_outlookmail_mail_id);
            list_outlookmail_mail_id.Clear();
            var folder = GetTargetFolder();
            foreach (object item in folder.Items)
            {
                if (item is MailItem mailItem)
                {
                    //if (ignoreFileList.Exists(mailItem.Parent.StoreID,mailItem.EntryID))    //here
                    if (ignoreFileList.Exists(folder.StoreID, mailItem.EntryID))   
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
                    //MailID mailID = new() { StoreID = mailItem.Parent.StoreID, EntryID = mailItem.EntryID };    //here
                    MailID mailID = new() { StoreID = folder.StoreID, EntryID = mailItem.EntryID };    //here
                    List_OutlookMailID.Add(mailID);
                    if (result.ActivateWindow == false && target.ActivateWindow && !original_list_outlookmail_mail_id.Contains(mailID))
                    {
                        result.ActivateWindow = true;
                    }
                }
            }
            result.List_new_mail_id = list_outlookmail_mail_id.Except(original_list_outlookmail_mail_id).ToList();
            return result;
        }
    }
}
