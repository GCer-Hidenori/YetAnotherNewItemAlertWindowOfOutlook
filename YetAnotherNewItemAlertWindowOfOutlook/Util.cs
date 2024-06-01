using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Util
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public static string ReadSettingSampleXmlString()
        {
            System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
            var info = System.Windows.Application.GetResourceStream(new Uri("setting.sample.xml", UriKind.Relative));
            using (var sr = new StreamReader(info.Stream))
            {
                return sr.ReadToEnd();
            }
        }
        private static MAPIFolder? GetSingleSearchFolder(Microsoft.Office.Interop.Outlook.Application outlook)
        {
            foreach (Store store in outlook.Session.Stores)
            {
                foreach (MAPIFolder folder in store.GetSearchFolders())
                {
                    return folder;
                }
            }
            return null;
        }

        public static Setting CreateInitialSettingFile(Microsoft.Office.Interop.Outlook.Application outlook, string settingFilePath)
        {
            Setting setting = new();

            var searchFolder = GetSingleSearchFolder(outlook);
            if (searchFolder != null)
            {
                var target_search_folder = new Target();
                target_search_folder.TargetFolderType = Target.FolderType.SearchFolder;
                target_search_folder.Path = searchFolder.FullFolderPath;
                setting.Targets.Add(target_search_folder);
            }
            var target_normal_folder = new Target();
            var inboxFolder = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            target_normal_folder.TargetFolderType = Target.FolderType.NormalFolder;
            target_normal_folder.Path = inboxFolder.FullFolderPath;
            setting.Targets.Add(target_normal_folder);

            var action1 = new Action()
            {
                ActionType = ActionType.ActivateWindow
            };
            var action2 = new Action()
            {
                ActionType = ActionType.CreateFile,
                FileName = @"c:\temp\a.txt",
                Body = "aaaaaaaa"
            };
            var filter_condition = new Condition()
            {
                Type = ConditionType.And,
                Conditions = new List<Condition>()
                {
                    new Condition()
                    {
                        Type = ConditionType.Subject,
                        Value = "test"
                    },
                    new Condition()
                    {
                        Type=ConditionType.SenderName,
                        Value = "test"
                    },
                    new Condition()
                    {
                        Type = ConditionType.SenderAddress,
                        Value = "aaa@example.com"
                    }
                }
            };
            target_normal_folder.Condition = filter_condition;

            target_normal_folder.Actions.Add(action1);
            target_normal_folder.Actions.Add(action2);
            setting.Save();
            return setting;
        }
    }
}
