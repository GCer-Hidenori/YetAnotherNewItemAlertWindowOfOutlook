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
            target_normal_folder.MailReceivedDaysThreshold = 30;
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

        //https://dobon.net/vb/dotnet/process/getprocessesbyfilename.html
        public static System.Diagnostics.Process[] GetProcessesByFileName(string searchFileName)
        {
            searchFileName = searchFileName.ToLower();
            System.Collections.ArrayList list = new System.Collections.ArrayList();

            //すべてのプロセスを列挙する
            foreach (System.Diagnostics.Process p
                in System.Diagnostics.Process.GetProcesses())
            {
                string fileName;
                try
                {
                    //メインモジュールのパスを取得する
                    fileName = p.MainModule.FileName;
                }
                catch (System.ComponentModel.Win32Exception)
                {
                    //MainModuleの取得に失敗
                    fileName = "";
                }
                if (0 < fileName.Length)
                {
                    //ファイル名の部分を取得する
                    fileName = System.IO.Path.GetFileName(fileName);
                    //探しているファイル名と一致した時、コレクションに追加
                    if (searchFileName.Equals(fileName.ToLower()))
                    {
                        list.Add(p);
                    }
                }
            }

            //コレクションを配列にして返す
            return (System.Diagnostics.Process[])
                list.ToArray(typeof(System.Diagnostics.Process));
        }
    }
}
