using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class OutlookUtil
    {
        public static MAPIFolder GetSearchFolder(string path)
        {
            var outlook = new Application();

            if (path.Length >= 2 && path.Substring(0, 2) == @"\\")
            {
                path = path.Substring(2, path.Length - 2);
            }
            string[] pathParts = path.Split('\\');
            if (pathParts.Length < 2)
            {
                throw new YError(ErrorType.InvalidTargetFolderPath, $"path: {path} ");
            }

            Store store;
            try
            {
                store = outlook.Session.Stores[pathParts[0]];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new YError(ErrorType.StoreNotFound, $"path:{path}");
            }
            foreach (MAPIFolder search_folder in store.GetSearchFolders())
            {
                if (search_folder.Name == pathParts[pathParts.Length - 1])
                {
                    return search_folder;
                }
            }
            throw new YError(ErrorType.NoFolderFoundError, path);

        }
        private static MAPIFolder GetNormalFolder(MAPIFolder parentFolder, List<string> child_folder_names)
        {
            if (child_folder_names.Count == 0)
            {
                return parentFolder;
            }
            foreach (MAPIFolder folder in parentFolder.Folders)
            {
                if (folder.Name == child_folder_names[0])
                {
                    child_folder_names.RemoveAt(0);
                    return GetNormalFolder(folder, child_folder_names);
                }
            }
            throw new YError(ErrorType.NoFolderFoundError, parentFolder.FullFolderPath + "/" + child_folder_names[0]);
        }
        public static MAPIFolder GetNormalFolder(string path)
        {
            var outlook = new Application();

            if (path.Length >= 2 && path.Substring(0, 2) == @"\\")
            {
                path = path.Substring(2, path.Length - 2);
            }
            List<string> pathParts = path.Split('\\').ToList<string>();
            if (pathParts.Count < 2)
            {
                throw new ArgumentException();
            }

            Store store;
            try
            {
                store = outlook.Session.Stores[pathParts[0]];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new YError(ErrorType.NoFolderFoundError, "store not found " + pathParts[0]);
            }
            pathParts.RemoveAt(0);
            return GetNormalFolder(store.GetRootFolder(), pathParts);
        }
        public static MailItem GetMail(string storeID, string entryID, Microsoft.Office.Interop.Outlook.Application outlook)
        {
            var ns = outlook.GetNamespace("MAPI");
            return ns.GetItemFromID(entryID, storeID);

        }
        public static void ListAllFolders(NLog.Logger logger)
        {
            logger.Info("list folders");
            ListRootFolders(logger);
            logger.Info("list search folders");
            ListSearchFolders(logger);
        }
        static void ListRootFolders(NLog.Logger logger)
        {
            var outlook = new Application();
            var ns = outlook.GetNamespace("MAPI");
            foreach (var folder in ns.Folders)
            {
                GetFolders((MAPIFolder)folder, logger);
            }
        }
        public static void GetFolders(MAPIFolder folder, NLog.Logger logger)
        {
            logger.Info(folder.FullFolderPath);
            foreach (MAPIFolder subFolder in folder.Folders)
            {
                GetFolders(subFolder, logger);
            }
        }
        static void ListSearchFolders(NLog.Logger logger)
        {
            var outlook = new Application();

            foreach (Store store in outlook.Session.Stores)
            {
                logger.Info(store.DisplayName);
                foreach (MAPIFolder folder in store.GetSearchFolders())
                {
                    logger.Info(folder.FullFolderPath);
                }
            }
        }
        public static MAPIFolder? GetSameThreadMailFolder(MailItem mailItem)
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");

            string folder_path = mailItem.Parent.FolderPath;
            string sent_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail).FolderPath;
            string draft_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderDrafts).FolderPath;
            string conflicts_folder_path = "";
            try
            {
                conflicts_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderConflicts).FolderPath;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                //
            }
            Conversation conversation = mailItem.GetConversation();

            //MailItem? sameThreadMailItem = conversation.GetRootItems().Cast<object>().FirstOrDefault(m => (TypeDescriptor.GetProperties(m)["MessageClass"].GetValue(m)== "IPM.Note\r\n" &&   m.Parent.FolderPath != folder_path && m.Parent.FolderPath != sent_folder_path && m.Parent.FolderPath != draft_folder_path && m.Parent.FolderPath != conflicts_folder_path));


            foreach (object object_samethread_root_mailItem in conversation.GetRootItems())
            {
                if (TypeDescriptor.GetProperties(object_samethread_root_mailItem)["MessageClass"].GetValue(object_samethread_root_mailItem).ToString() == "IPM.Note")
                {
                    MailItem samethread_root_mailItem = (MailItem)object_samethread_root_mailItem;
                    MAPIFolder samethread_root_mailItem_folder = samethread_root_mailItem.Parent;
                    if (samethread_root_mailItem_folder.FolderPath != folder_path && samethread_root_mailItem_folder.FolderPath != sent_folder_path && samethread_root_mailItem_folder.FolderPath != draft_folder_path && samethread_root_mailItem_folder.FolderPath != conflicts_folder_path)
                    {
                        return samethread_root_mailItem.Parent;
                    }
                    //while (Marshal.ReleaseComObject(samethread_root_mailItem) > 0) { }
                    samethread_root_mailItem = null;

                }
                foreach (object object_samethread_mailItem in conversation.GetChildren(object_samethread_root_mailItem))
                {
                    if (TypeDescriptor.GetProperties(object_samethread_mailItem)["MessageClass"].GetValue(object_samethread_mailItem).ToString() == "IPM.Note")
                    {
                        MailItem samethread_mailItem = (MailItem)object_samethread_mailItem;
                        if (samethread_mailItem.Parent.FolderPath != folder_path && samethread_mailItem.Parent.FolderPath != sent_folder_path && samethread_mailItem.Parent.FolderPath != draft_folder_path && samethread_mailItem.Parent.FolderPath != conflicts_folder_path)
                        {
                            return samethread_mailItem.Parent;
                        }
                        //while (Marshal.ReleaseComObject(samethread_mailItem) > 0) { }
                        samethread_mailItem = null;

                    }
                }
                //while (Marshal.ReleaseComObject(object_samethread_root_mailItem) > 0) { }

            }
            return null;
        }
        public static List<MailItem> GetSameFolderSameThreadEachMails(MailItem mailItem)
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            var sameFolderSameThreadMails = new List<MailItem>();

            string folder_path = mailItem.Parent.FolderPath;
            string sent_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail).FolderPath;
            string draft_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderDrafts).FolderPath;
            string conflicts_folder_path = "";
            try
            {
                conflicts_folder_path = ns.GetDefaultFolder(OlDefaultFolders.olFolderConflicts).FolderPath;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                //
            }
            Conversation conversation = mailItem.GetConversation();

            foreach (object object_samethread_root_mailItem in conversation.GetRootItems())
            {
                foreach (object object_samethread_mailItem in conversation.GetChildren(object_samethread_root_mailItem))
                {
                    if (TypeDescriptor.GetProperties(object_samethread_mailItem)["MessageClass"].GetValue(object_samethread_mailItem).ToString() == "IPM.Note")
                    {
                        MailItem samethread_mailItem = (MailItem)object_samethread_mailItem;
                        if (samethread_mailItem.Parent.FolderPath == folder_path)
                        {
                            sameFolderSameThreadMails.Add(samethread_mailItem);
                        }
                    }
                }
            }
            return sameFolderSameThreadMails;
        
        }
        public static List<MailItem> GetSameFolderSameThreadMails(List<MailItem> mails)
        {
            var sameFolderSameThreadMails = new List<MailItem>();
            foreach (MailItem mailItem in mails)
            {

                foreach (var sameFolderSameThreadEachMail in GetSameFolderSameThreadEachMails(mailItem))
                {
                    sameFolderSameThreadMails.Add(sameFolderSameThreadEachMail);
                }

            }
            return sameFolderSameThreadMails;
        }
        public static List<MailItem> Items2MailItems(Items mails)
        {
            var mailItems = new List<MailItem>();
            foreach (object item in mails)
            {
                if (item is MailItem mailItem)
                {
                    mailItems.Add(mailItem);
                }
            }
            return mailItems;
        }
    }

}
