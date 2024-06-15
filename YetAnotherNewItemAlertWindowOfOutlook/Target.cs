//using System.Xml;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
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

        private Filter? filter = null;
        private int timers_to_check_mail = 1;   //How many timers does it take to start?
        private FolderType folderType;
        private string path = "";
        private int? mailReceivedDaysThreshold = null;
        private bool viewSameThreadSameFolderMail = false;

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
        public int? MailReceivedDaysThreshold { get => mailReceivedDaysThreshold; set => mailReceivedDaysThreshold = value; }

        public bool ViewSameThreadSameFolderMail { get => viewSameThreadSameFolderMail; set => viewSameThreadSameFolderMail = value; }

        //private bool activateWindow = false;


        private List<Rule> rules = new();

        /*
        public bool ActivateWindow
        {
            get
            {
                return Actions.Any(x => x.ActionType == ActionType.ActivateWindow);
            }
        }
        */




        //internal List<Action> ActionCreateFiles { get => Actions; set => Actions = value; }
        public Filter? Filter { get => filter; set => filter = value; }
        [XmlArray("Rules")]
        public List<Rule> Rules { get => rules; set => rules = value; }

        public bool Filtering(MailItem mailItem)
        {
            if (Filter != null)
            {
                return Filter.Evaluate(mailItem);
            }
            else
            {
                return true;
            }
        }
    }

}
