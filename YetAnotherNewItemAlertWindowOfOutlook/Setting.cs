using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using YetAnotherNewItemAlertWindowOfOutlook.Properties;
using System.Xml;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Setting
    {
        private List<Target> targets = new();
        private static readonly string fileName = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "setting.xml");
        private int timer_interval_sec = 60;

        public List<Target> Targets
        {
            get { return targets; }
            set { targets = value; }
        }

        public int TimerIntervalSec {
            get => timer_interval_sec; 
            set
            {
                if (value > 0)
                {
                    timer_interval_sec = value;
                }
                else
                {
                    timer_interval_sec = 60;
                }
            }
        }

        /*
        public void Save()
        {
            System.Xml.Serialization.XmlSerializer serializer = new(typeof(Setting));
            using(System.IO.StreamWriter sw = new(fileName, false, new System.Text.UTF8Encoding(false))){
	            serializer.Serialize(sw, this);
            }
        }
        */

        public static Setting Load(NLog.Logger logger)
        {
            XmlDocument xdoc = new();
            if (File.Exists(fileName))
            {
                try
                {
                    xdoc.Load(fileName);
                    var setting = new Setting();
                    XmlNode x;
                    if((x=xdoc.SelectSingleNode("/Setting/Timer[@interval_sec!='']")) != null)
                    {
                        setting.TimerIntervalSec = int.Parse(x.Attributes["interval_sec"].Value);
                    }
               
                    xdoc.SelectNodes("/Setting/Targets/Target").Cast<XmlNode>().ToList().ForEach(x =>
                    {
                        var target = new Target();
                        target.Logger = logger;
                        target.IntervalMin = int.Parse(x.SelectSingleNode("IntervalMin").InnerText);
                        target.TargetFolderType = (Target.FolderType)Enum.Parse(typeof(Target.FolderType), x.SelectSingleNode("TargetFolderType").InnerText);
                        target.Path = x.SelectSingleNode("Path").InnerText;
                        XmlNode x2;
                        if ((x2 = x.SelectSingleNode("Actions/Activate_Window")) != null)
                        {
                            if (x2.InnerText.ToUpper() == "TRUE")
                            {
                                target.ActivateWindow = true;
                            }
                        }
                        foreach (XmlNode xCreateFile in x.SelectNodes("Actions/Create_File"))
                        {
                            var createFile = new ActionCreateFile(logger);
                            if (xCreateFile.Attributes["fileName"] != null)
                            {
                                createFile.FileName = xCreateFile.Attributes["fileName"].Value;
                            }
                            else
                            {
                                throw new YError(ErrorType.ActionCreateFileError, "Create_File element must have fileName attribute.");
                            }
                            XmlNode xBody;
                            if ((xBody = xCreateFile.SelectSingleNode("body")) != null)
                            {
                                createFile.Body = xBody.InnerText;
                            }
                            target.ActionCreateFiles.Add(createFile);

                        }
                        if (x.SelectSingleNode("Filter") != null)
                        {
                            target.FilterNode = x.SelectSingleNode("Filter");
                        }
                        setting.Targets.Add(target);
                    });
                    return setting;
                }
                catch (XmlException e)
                {
                    string message = $@"source:{e.Source}
Message:{e.Message}
Line number:{e.LineNumber}
Line position:{e.LinePosition}
";
                    throw new YError(ErrorType.SettingFileLoadError,message);

                }
            }
            else
            {
                throw new YError(ErrorType.SettingFileNotFound,fileName);
            }
        }
        /*
        public static Setting Init(){
            if (System.IO.File.Exists(fileName))
            {
                return Load();
            }
            else{
                Setting setting = new();
                setting.Save();
                return setting;
            }
        }
        */
    }
}
