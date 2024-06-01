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
        private string path = "";
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
        public String Path
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
    }

}
