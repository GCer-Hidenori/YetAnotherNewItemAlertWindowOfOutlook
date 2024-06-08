using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public enum ActionType
    {
        ActivateWindow,
        AddCategory,
        CreateFile
    }
    [XmlRoot("Action")]
    public class Action
    {
        private ActionType action_type;
        private string? fileName = null;
        private string? body = null;
        private string? attribute_value = null;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        [XmlAttribute("fileName")]
        public string? FileName { get => fileName; set => fileName = value; }

        [XmlElement("body")]
        public string? Body { get => body; set => body = value; }

        [XmlAttribute("type")]
        public ActionType ActionType { get => action_type; set => action_type = value; }
        [XmlAttribute("value")]
        public string? AttributeValue { get => attribute_value; set => attribute_value = value; }


        private string ReplaceWords(string? source, MailItem mailItem)
        {

            source = Regex.Replace(source ?? "", @"\${entryID}", mailItem.EntryID, RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${Subject}", mailItem.Subject, RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${ReceivedTime}", mailItem.ReceivedTime.ToString(), RegexOptions.IgnoreCase).Replace('/', '-');
            source = Regex.Replace(source, @"\${SenderName}", mailItem.SenderName ?? "", RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${SenderEmailAddress}", mailItem.SenderEmailAddress ?? "", RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${Body}", mailItem.Body, RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${To}", mailItem.To ?? "", RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${Cc}", mailItem.CC ?? "", RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${Categories}", mailItem.Categories ?? "", RegexOptions.IgnoreCase);
            source = Regex.Replace(source, @"\${SentOn}", mailItem.SentOn.ToString(), RegexOptions.IgnoreCase).Replace('/', '-');
            return source;
        }
        public void CreateFile(MailItem mailItem)
        {
            string valfileName = ReplaceWords(fileName, mailItem);
            string valBody = ReplaceWords(body, mailItem);

            Encoding enc = Encoding.UTF8;
            using (StreamWriter writer = new StreamWriter(valfileName, false, enc))
            {
                writer.WriteLine(valBody);
            }
            Logger.Info($"file create or updated {valfileName}");
        }
        public void AddCategory(MailItem mailItem,string? categoryID)
        {
            if(categoryID != null)
            {
                string registry_key_name = @"HKEY_CURRENT_USER\Control Panel\International";
                string registry_value_nake = "sList";
                RegistryKey registry_key = Registry.LocalMachine.OpenSubKey(registry_key_name);

                // レジストリの値を取得
                string delimiter = (string)registry_key.GetValue(registry_value_nake);

                if ( new List<string>(mailItem.Categories.Split(delimiter)).Exists(c => c.ToUpper()==categoryID.ToUpper() ))
                {
                    return;
                }

                mailItem.Categories = mailItem.Categories + delimiter + categoryID;
                mailItem.Save();
            }

        }

        public void Execute(MailItem mailItem, Window window)
        {
            switch (action_type)
            {
                case ActionType.CreateFile:
                    CreateFile(mailItem);
                    break;
                case ActionType.ActivateWindow:
                    Logger.Info("activate window.");
                    window.Dispatcher.Invoke(() =>
                        {
                            window.Activate();
                            window.WindowState = WindowState.Normal;
                        });
                    break;
                case ActionType.AddCategory:
                    Logger.Info("add category.");
                    AddCategory(mailItem, attribute_value);
                    break;
            }
        }
    }
}
