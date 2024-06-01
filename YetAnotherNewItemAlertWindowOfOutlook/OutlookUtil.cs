using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;

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
    }

}
