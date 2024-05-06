using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Xml;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Util
    {
        private static System.Xml.XmlDocument ReadSettingSampleXml(NLog.Logger logger)
        {
            System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
            string xmlString = ReadSettingSampleXmlString();
            try
            {
                xdoc.LoadXml(xmlString);
            }
            catch (XmlException e)
            {
                    string message = $@"source:{e.Source}
Message:{e.Message}
Line number:{e.LineNumber}
Line position:{e.LinePosition}
xml:{xmlString}
";
                    throw new YError(ErrorType.SampleSettingFileLoadError,message);
            }
            return xdoc;
        }
        public static string ReadSettingSampleXmlString()
        {
            System.Xml.XmlDocument xdoc = new System.Xml.XmlDocument();
            var info = System.Windows.Application.GetResourceStream(new Uri("setting.sample.xml", UriKind.Relative));
            using (var sr = new StreamReader(info.Stream))
            {
                return sr.ReadToEnd();
            }
        }
        private static MAPIFolder GetSingleSearchFolder(Microsoft.Office.Interop.Outlook.Application outlook)
        {   
            foreach(Store store in outlook.Session.Stores)
            {
                foreach(MAPIFolder folder in store.GetSearchFolders())
                {
                    return folder;
                }
            }
            return null;
        }
        public static void CreateSettingFile(Microsoft.Office.Interop.Outlook.Application outlook, string settingFilePath,NLog.Logger logger)
        {
            var xdoc = ReadSettingSampleXml(logger);
            var xTarget_NormalFolder = xdoc.SelectSingleNode("//Target[TargetFolderType='NormalFolder']");
            var inboxFolder = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            xTarget_NormalFolder.SelectSingleNode(".//Path").InnerText = inboxFolder.FolderPath;

            var searchFolder = GetSingleSearchFolder(outlook);
            if(searchFolder != null)
            {
                var xTarget_SearchFolder = xdoc.SelectSingleNode("//Target[TargetFolderType='SearchFolder']");
                xTarget_SearchFolder.SelectSingleNode(".//Path").InnerText = searchFolder.FolderPath;
            }
            else
            {
                xdoc.SelectSingleNode("//Targets").RemoveChild(xdoc.SelectSingleNode("//Target[TargetFolderType='SearchFolder']"));
            }

            xdoc.Save(settingFilePath);
        }
    }
}
