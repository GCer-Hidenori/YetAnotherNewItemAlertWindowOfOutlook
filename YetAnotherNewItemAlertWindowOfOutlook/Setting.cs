using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using YetAnotherNewItemAlertWindowOfOutlook.Properties;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Setting
    {
        private List<Target> targets = new();
        private static readonly string fileName = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "setting.xml");
        private int timer_interval_sec = 60;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();


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

        
        public void Save()
        {
            System.Xml.Serialization.XmlSerializer serializer = new(typeof(Setting));
            using(System.IO.StreamWriter sw = new(fileName, false, new System.Text.UTF8Encoding(false))){
	            serializer.Serialize(sw, this);
            }
        }
        

        public static Setting Load()
        {
            Setting setting = null;
            XmlSerializer serializer = new XmlSerializer(typeof(Setting));
            using (var sr = new StreamReader(fileName))
            {
                setting = (Setting)serializer.Deserialize(sr);
            }
            return setting;
        }


      
    }
}
