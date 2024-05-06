using System;
using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class IgnoreFile
    {
        private static readonly string ignoreDirPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "excluded_ids");

        public static bool Exists(string entryID)
        {
            if (Directory.Exists(ignoreDirPath))
            {
                return File.Exists(Path.Combine(ignoreDirPath, entryID));
            }
            return false;
        }
        public static void Add(string entryID,OutlookMailItem outlookMailitem,NLog.Logger logger)
        {
            if (!Directory.Exists(ignoreDirPath))
            {
                Directory.CreateDirectory(ignoreDirPath);
            }
            string fileName = Path.Combine(ignoreDirPath, entryID);
            string body = "subject: " + outlookMailitem.Subject + "\n" + "receivedTime: " + outlookMailitem.ReceivedTime + "\n" + "from: " + outlookMailitem.SenderName + "<" + outlookMailitem.SenderEmailAddress + ">"; 

            Encoding enc = Encoding.UTF8;
            using (StreamWriter writer = new StreamWriter(fileName, false, enc))
            {
                writer.WriteLine(body);
            }
            logger.Info($"ignore file create or updated {fileName}");
        }
    }
}
