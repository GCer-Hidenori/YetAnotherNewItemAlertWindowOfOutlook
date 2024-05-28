using System.IO;
using System.Text;
using Microsoft.Office.Interop.Outlook;

/*
namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class IgnoreFile
    {
        private static readonly string ignoreDirPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "excluded_ids");
        public static void ClearIgnoreList()
        {
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");
            foreach(var file in Directory.GetFiles(ignoreDirPath))
            {
                var entryID = System.IO.Path.GetFileName(file);
                var mailItem = ns.GetItemFromID(entryID);
                
            }
        }
        public static string GetIgnoreListDir()
        {
            return ignoreDirPath;
        }
        public static bool Exists(string storeID,string entryID)
        {
            if (Directory.Exists(ignoreDirPath))
            {
                return File.Exists(Path.Combine(ignoreDirPath,storeID+"_" + entryID));
            }
            return false;
        }
        public static void Add(string storeID, string entryID, OutlookMailItem outlookMailitem, NLog.Logger logger)
        {
            if (!Directory.Exists(ignoreDirPath))
            {
                Directory.CreateDirectory(ignoreDirPath);
            }
            string fileName = Path.Combine(ignoreDirPath,storeID+"_"+ entryID);
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
*/