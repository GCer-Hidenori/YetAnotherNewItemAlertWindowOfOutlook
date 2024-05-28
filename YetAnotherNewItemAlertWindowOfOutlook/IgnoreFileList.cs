using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using YetAnotherNewItemAlertWindowOfOutlook.Properties;
using System.Runtime.Serialization;
using System.Collections;


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

        public bool Exists(string storeID,string entryID)
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
        public void Add(string storeID,string entryID)
        {
            if (!ignoreFileList.ContainsKey(storeID))
            {
                ignoreFileList[storeID] = new HashSet<string>();
            }
            ignoreFileList[storeID].Add(entryID);
        }

    }
}
