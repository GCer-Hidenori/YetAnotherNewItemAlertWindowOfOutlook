using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public enum ActionType
    {
        ActivateWindow,
        CreateFile
    }
    [XmlRoot("Action")]
    public class Action
    {
        private ActionType action_type;
        private string? fileName = null;
        private string? body = null;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        [XmlAttribute("fileName")]
        public string? FileName { get => fileName; set => fileName = value; }

        [XmlElement("body")]
        public string? Body { get => body; set => body = value; }

        [XmlAttribute("type")]
        public ActionType ActionType { get => action_type; set => action_type = value; }


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
    }
}
