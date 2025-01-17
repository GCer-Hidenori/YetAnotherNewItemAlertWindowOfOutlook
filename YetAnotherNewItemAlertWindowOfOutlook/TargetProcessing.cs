using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

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

        private Items GetItems(MAPIFolder folder)
        {
            if(target.MailReceivedDaysThreshold == null)
            {
                return folder.Items;
            }
            else
            {
                string filter = "[ReceivedTime] >= '" + System.DateTime.Now.AddDays(-1 * target.MailReceivedDaysThreshold.Value).ToString("yyyy/MM/dd HH:mm") + "'";
                return folder.Items.Restrict(filter);
            }
        }

        public ResultOfTargetProcessing RefreshOutlookMailItem(IgnoreFileList ignoreFileList, Window window)
        {
            var result = new ResultOfTargetProcessing();
            List<MailID> original_list_outlookmail_mail_id = new(list_outlookmail_mail_id);
            list_outlookmail_mail_id.Clear();
            var folder = GetTargetFolder();
            var listMailItems = OutlookUtil.Items2MailItems(GetItems(folder));
            /* Issue #5 Outlook mail search takes too long
            if(target.TargetFolderType==Target.FolderType.SearchFolder &&  target.ViewSameThreadSameFolderMail)
            {
                var sameFolderSameThreadMails = OutlookUtil.GetSameFolderSameThreadMails(listMailItems);
                listMailItems.AddRange(sameFolderSameThreadMails);
                listMailItems = listMailItems.Distinct().ToList();
            }
            */

            foreach (MailItem mailItem in listMailItems)
            {

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
                MailID mailID = new() { StoreID = folder.StoreID, EntryID = mailItem.EntryID };    //here
                List_OutlookMailID.Add(mailID);
                while (Marshal.ReleaseComObject(mailItem) > 0) { }

            }
            result.List_new_mail_id = list_outlookmail_mail_id.Except(original_list_outlookmail_mail_id).ToList();
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");

            foreach (MailID new_mail_id in result.List_new_mail_id)
            {
                foreach (Rule rule in target.Rules)
                {
                    MailItem mailItem;
                    try
                    {
                        mailItem = ns.GetItemFromID(new_mail_id.EntryID, new_mail_id.StoreID);
                        if (rule.Condition != null && rule.Condition.Evaluate(mailItem))
                        {
                            foreach (Action action in rule.Actions)
                            {
                                action.Execute(mailItem, window);
                            }
                        }
                        while (Marshal.ReleaseComObject(mailItem) > 0) { }
                        mailItem = null;
                    }catch (System.Runtime.InteropServices.COMException e)
                    {
                        Logger.Error("Error happened when processing mail item. Skip this mail item.");
                        Logger.Error($"mail entry id {new_mail_id.EntryID}");
                        Logger.Error(e);
                    }

                }
            }
            return result;
        }
    }
}
