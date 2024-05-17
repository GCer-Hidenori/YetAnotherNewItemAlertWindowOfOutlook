//using System.Xml;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Target
    {
        public enum FolderType
        {
            NormalFolder,
            SearchFolder
        }

        private Condition? condition = null;
        private int timers_to_check_mail = 1;   //How many timers does it take to start?
        private FolderType folderType;
        private string? path;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public FolderType TargetFolderType
        {
            get { return folderType; }
            set { folderType = value; }
        }
        public int TimersToCheckMail
        {
            get { return timers_to_check_mail; }
            set { timers_to_check_mail = value; }
        }
        public String? Path
        {
            get { return path; }
            set { path = value; }
        }
        //private bool activateWindow = false;


        private List<Action> actions = new();

        public bool ActivateWindow
        {
            get
            {
                return Actions.Any(x => x.ActionType == ActionType.ActivateWindow);
            }
        }




        //internal List<Action> ActionCreateFiles { get => Actions; set => Actions = value; }
        public Condition? Condition { get => condition; set => condition = value; }
        [XmlArray("Actions")]
        public List<Action> Actions { get => actions; set => actions = value; }

        public bool Filtering(MailItem mailItem)
        {
            if (Condition != null)
            {
                return Condition.Evaluate(mailItem);
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
            }else  if (element.Name.ToUpper() == "And")
            {
                return element.ChildNodes.Cast<XmlNode>().Where(x => x.NodeType == XmlNodeType.Element).All(x => Filtering(mailItem, x));
            }
            else if (element.Name.ToUpper() == "Or")
            {
                return element.ChildNodes.Cast<XmlNode>().Where(x => x.NodeType == XmlNodeType.Element).Any(x => Filtering(mailItem, x));
            }
            else if (element.Name.ToUpper() == "Not")
            {
                return !Filtering(mailItem, element.FirstChild);
            }
            else if (element.Name.ToUpper() == "Subject")
            {
                return mailItem.Subject.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "Body")
            {
                return mailItem.Body.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "SenderAddress")
            {
                return mailItem.SenderEmailAddress.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "SenderName")
            {
                return mailItem.SenderName.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "RecipientNames")
            {
                return String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Name)).Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "RECIPIENTADDRESSES")
            {
                return String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Address)).Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "To")
            {
                return mailItem.To.Contains(element.InnerText);
            }
            else if (element.Name.ToUpper() == "Cc")
            {
                return mailItem.Cc.Contains(element.InnerText);
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
