using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class ActionCreateFile
    {
        private string fileName;
        private string body = "";
        private NLog.Logger logger;

        public string FileName { get => fileName; set => fileName = value; }
        public string Body { get => body; set => body = value; }

        public ActionCreateFile(NLog.Logger logger)
        {
            this.logger = logger;
        }
        private string ReplaceWords(string source,MailItem mailItem)
        {
            
            source =  Regex.Replace(source, @"\${entryID}", mailItem.EntryID, RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${Subject}", mailItem.Subject, RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${ReceivedTime}", mailItem.ReceivedTime.ToString(), RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${SenderName}", mailItem.SenderName ?? "", RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${SenderEmailAddress}", mailItem.SenderEmailAddress ?? "", RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${Body}", mailItem.Body, RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${To}", mailItem.To ?? "", RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${Cc}", mailItem.CC ?? "", RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${Categories}", mailItem.Categories ?? "", RegexOptions.IgnoreCase);
            source =  Regex.Replace(source, @"\${SentOn}", mailItem.SentOn.ToString(), RegexOptions.IgnoreCase);
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
            logger.Info($"file create or updated {valfileName}");
        }
    }
}
