using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Setting
    {
        private List<Target> targets = new();
        public static readonly string fileName = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "setting.xml");
        private int timer_interval_sec = 60;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        private List<Column> columns = new();

        public int TimerIntervalSec
        {
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

        public List<Target> Targets
        {
            get { return targets; }
            set { targets = value; }
        }


        public List<Column> Columns { get => columns; set => columns = value; }

        public void Save()
        {
            System.Xml.Serialization.XmlSerializer serializer = new(typeof(Setting));
            using (System.IO.StreamWriter sw = new(fileName, false, new System.Text.UTF8Encoding(false)))
            {
                serializer.Serialize(sw, this);
            }
        }

        public static Setting Init()
        {
            if (File.Exists(fileName))
            {
                return Load();
            }
            else
            {
                var outlook = new Microsoft.Office.Interop.Outlook.Application();
                return Util.CreateInitialSettingFile(outlook, fileName);
            }
        }

        private static Setting Load()
        {
            Setting? setting;
            XmlSerializer serializer = new XmlSerializer(typeof(Setting));
            using (var sr = new StreamReader(fileName))
            {
                setting = (Setting?)serializer.Deserialize(sr);
            }
            if (setting != null)
            {
                return setting;
            }
            else
            {
                throw new YError(ErrorType.SettingFileLoadError);
            }
        }



    }
}
