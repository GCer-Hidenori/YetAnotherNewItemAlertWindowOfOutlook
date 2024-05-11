using System;
using System.Collections.Generic;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
//using System.Xml;
using Microsoft.Office.Interop.Outlook;
using NLog;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Target
    {
        public enum FolderType
        {
            NormalFolder,
            SearchFolder
        }

        //private XmlNode? filterNode = null;   
        private Condition? filter = null;
        private int interval_min;
        private FolderType folderType;
        private string? path;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public FolderType TargetFolderType
        {
            get { return folderType; }
            set { folderType = value; }
        }
        public int IntervalMin
        {
            get { return interval_min; }
            set { interval_min = value; }
        }
        public String? Path
        {
            get { return path; }
            set { path = value; }
        }
        private bool activateWindow = false;
        private List<ActionCreateFile> actionCreateFiles = new();
        public bool ActivateWindow { get => activateWindow; set => activateWindow = value; }

        //public XmlNode? FilterNode { get => filterNode; set => filterNode = value; }



        internal List<ActionCreateFile> ActionCreateFiles { get => actionCreateFiles; set => actionCreateFiles = value; }
        public Condition? Filter { get => filter; set => filter = value; }

        public bool Filtering(MailItem mailItem)
        {
            if(Filter != null)
            {
                return Filter.Evaluate(mailItem);
            }
            else
            {
                return true;
            }
        }
        /*
        public bool Filtering(MailItem mailItem, XmlNode? element)
        {
            if(element == null)
            {
                return true;
            }else  if (element.Name.ToUpper() == "AND")
            {
                return element.ChildNodes.Cast<XmlNode>().Where(x => x.NodeType == XmlNodeType.Element).All(x => Filtering(mailItem, x));
            }
            else if (element.Name.ToUpper() == "OR")
            {
                return element.ChildNodes.Cast<XmlNode>().Where(x => x.NodeType == XmlNodeType.Element).Any(x => Filtering(mailItem, x));
            }
            else if (element.Name.ToUpper() == "NOT")
            {
                return !Filtering(mailItem, element.FirstChild);
            }
            else if (element.Name.ToUpper() == "SUBJECT")
            {
                return mailItem.Subject.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "BODY")
            {
                return mailItem.Body.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "SENDEREMAILADDRESS")
            {
                return mailItem.SenderEmailAddress.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "SENDERNAME")
            {
                return mailItem.SenderName.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "RECIPIENTNAMES")
            {
                return String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Name)).Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "RECIPIENTADDRESSES")
            {
                return String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Address)).Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "TO")
            {
                return mailItem.To.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "CC")
            {
                return mailItem.CC.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "ATTACHMENT")
            {
                return mailItem.Attachments.Cast<Attachment>().Any(x => x.FileName.Contains(element.InnerText));
            }
            else
            {
                Logger.Error($"Invalid Filter Element Name: {element.Name}");
                throw new YError(ErrorType.InvalidFilterElementName,element.Name);
            }
        }
        */
    }
        
}
