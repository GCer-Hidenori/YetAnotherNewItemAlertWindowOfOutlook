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
                var target_search_folder = new Target
                {
                    TargetFolderType = Target.FolderType.SearchFolder,
                    Path = searchFolder.FullFolderPath,
                    ViewSameThreadSameFolderMail = true
                };
                setting.Targets.Add(target_search_folder);
            }
            var target_normal_folder = new Target
            {
                MailReceivedDaysThreshold = 30
            };
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


            var rule_condition = new Condition()
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
                        Value = "bbb@example.com"
                    }
                }
            };
            var rule1 = new Rule()
            {
                Condition = rule_condition,
                Actions = new List<Action>() { action1, action2 }
            };
            target_normal_folder.Rules.Add(rule1);

            var rule_condition2 = new Condition()
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
                        Value = "ccc@example.com"
                    }
                }
            };
            var rule2 = new Rule()
            {
                Condition = rule_condition2,
                Actions = new List<Action>() {
                    new Action(){
                        ActionType = ActionType.AddCategory,
                        AttributeValue = "important"
                    },
                    new Action(){
                        ActionType = ActionType.MoveMail,
                        AttributeValue = @"aaa\Inbox\Test"
                    }
                }
            };
            target_normal_folder.Rules.Add(rule2);


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
            var filter = new Filter() { Conditions = new List<Condition>() { filter_condition } };
            target_normal_folder.Filter = filter;

            setting.Save();
            return setting;
        }

     
    }
}
