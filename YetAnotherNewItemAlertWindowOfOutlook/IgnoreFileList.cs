using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;


namespace YetAnotherNewItemAlertWindowOfOutlook
{
    [DataContract]


    public class IgnoreFileList
    {
        [DataMember]
        public Dictionary<string, HashSet<string>> ignoreFileList = new();


        public readonly static string ignore_file_list_path = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "excluded_ids.xml");

        public static IgnoreFileList Load()
        {
            IgnoreFileList ignoreFileList;
            DataContractSerializer serializer = new DataContractSerializer(typeof(IgnoreFileList));
            using (FileStream fs = new FileStream(IgnoreFileList.ignore_file_list_path, FileMode.Open))
            {
                ignoreFileList = (IgnoreFileList)serializer.ReadObject(fs);
            }
            return ignoreFileList;

        }

        public void Save()
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(IgnoreFileList));
            using (FileStream fs = new FileStream(IgnoreFileList.ignore_file_list_path, FileMode.Create))
            {
                serializer.WriteObject(fs, this);
            }
        }

        public bool Exists(string storeID, string entryID)
        {
            if (ignoreFileList.ContainsKey(storeID))
            {
                return ignoreFileList[storeID].Contains(entryID);
            }
            else
            {
                return false;
            }
        }

        public static IgnoreFileList Init()
        {
            IgnoreFileList ignoreFileList;
            if (File.Exists(ignore_file_list_path))
            {
                ignoreFileList = IgnoreFileList.Load();
            }
            else
            {
                ignoreFileList = new IgnoreFileList();
            }
            return ignoreFileList;
        }
        public void Add(string storeID, string entryID)
        {
            if (!ignoreFileList.ContainsKey(storeID))
            {
                ignoreFileList[storeID] = new HashSet<string>();
            }
            ignoreFileList[storeID].Add(entryID);
        }
        public void DeleteUnwantedIgnoreList()
        {
            var outlook = new Application();
            var ns = outlook.GetNamespace("MAPI");

            var list = new List<string>();
            foreach (var storeID in ignoreFileList.Keys)
            {
                foreach (var entryID in ignoreFileList[storeID])
                {
                    try
                    {
                        MailItem mailItem = ns.GetItemFromID(entryID, storeID);
                        while (Marshal.ReleaseComObject(mailItem) > 0) { }
                        mailItem = null;
                    }
                    catch (System.Exception e)
                    {
                        ignoreFileList[storeID].Remove(entryID);
                    }
                }
            }
        }

    }
}
