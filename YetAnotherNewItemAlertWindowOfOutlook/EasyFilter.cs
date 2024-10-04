using Microsoft.Office.Interop.Outlook;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System.Xml;
using System.Windows.Automation;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public partial class MainWindow : Window
    {
        // private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();


        private void EasyFilter_DeleteMail(MailItem mailItem)
        {

            try
            {
                context?.HideMail(mailItem.EntryID, mailItem.Parent.StoreID);
                mailItem.Delete();
            }
            catch (System.Runtime.InteropServices.COMException e3)
            {
                MessageBox.Show("Can't delete mail.");
                Logger.Warn(e3);
            }
        }
        private bool EasyFilter_ConditionEvaluate(XmlElement conditionElement, MailItem mailItem)
        {
            switch (conditionElement.Name)
            {
                case "true":
                    return true;
                case "false":
                    return false;
                case "and":
                    return conditionElement.ChildNodes.Cast<XmlNode>().Where(n => n.NodeType == XmlNodeType.Element).Cast<XmlElement>().ToList().All(c => EasyFilter_ConditionEvaluate(c, mailItem));
                case "or":
                        return conditionElement.ChildNodes.Cast<XmlNode>().Where(n => n.NodeType == XmlNodeType.Element).Cast<XmlElement>().ToList().Any(c => EasyFilter_ConditionEvaluate(c, mailItem));
                case "not":
                    return !EasyFilter_ConditionEvaluate((XmlElement)conditionElement.FirstChild, mailItem);
                case "condition":
                    return EasyFilter_ConditionEvaluate((XmlElement)conditionElement.FirstChild, mailItem);
                case "subject":

                    if (mailItem.Subject.Contains(conditionElement.InnerText))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case "body":

                    if (mailItem.Body.Contains(conditionElement.InnerText))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case "from":
                    if (mailItem.SenderName.Contains(conditionElement.InnerText))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case "to":
                    if (mailItem.To.Contains(conditionElement.InnerText))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case "cc":
                    if (mailItem.CC?.Contains(conditionElement.InnerText) == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case "receiver":
                    if (mailItem.To.Contains(conditionElement.InnerText))
                    {
                        return true;
                    }
                    else if (mailItem.CC?.Contains(conditionElement.InnerText) == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                default:
                    return false;
            }


        }

        private void EasyFilter()
        {
            if(setting.EasyFilterXmlPath == null)
            {
                MessageBox.Show("Please set the EasyFilterXmlPath in the setting.");
                return;
            }
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            var ns = outlook.GetNamespace("MAPI");

            if (datagrid.SelectedItems.Count > 10)
            {
                MessageBox.Show("Too many items selected.");
                return;
            }

            var xmlDoc = new XmlDocument();
            try
            {
                
                xmlDoc.Load(setting.EasyFilterXmlPath);
            }catch(XmlException e)
            {
                Logger.Error($"can't open xml. {setting.EasyFilterXmlPath} ");
                Logger.Error(e);
                return;
            }
            
            List<OutlookMailItem> listSelectedItems = datagrid.SelectedItems.Cast<OutlookMailItem>().ToList();
            foreach (OutlookMailItem outlookMailItem in listSelectedItems)
            {
                EasyFilter_Operation(outlookMailItem,ns,xmlDoc);
            }
        }
        private void EasyFilter_Operation(OutlookMailItem outlookMailItem,NameSpace ns,XmlDocument xmlDoc)
        {
            XmlNode root = xmlDoc.DocumentElement;
            MailItem mailItem;
            try
            {
                mailItem = ns.GetItemFromID(outlookMailItem.EntryID, outlookMailItem.StoreID);
            }
            catch (System.Runtime.InteropServices.COMException e2)
            {
                MessageBox.Show("Can't open mail.");
                Logger.Warn(e2);
                return;
            }
            foreach (XmlElement filterElement in root.ChildNodes.Cast<XmlNode>().Where(n => n.NodeType == XmlNodeType.Element))
            {
                XmlElement conditionElement = (XmlElement)filterElement.SelectSingleNode("condition");
                if (EasyFilter_ConditionEvaluate(conditionElement, mailItem))
                {
                    foreach (XmlElement operationElement in filterElement.SelectNodes("operation/*"))
                    {
                        switch (operationElement.Name)
                        {
                            case "delete":
                                if (operationElement.GetAttribute("kakunin") != null)
                                {
                                    MessageBoxResult res2 = MessageBox.Show($"Would you like to delete this mail from Outlook?", "Confirmation", MessageBoxButton.YesNoCancel);
                                    switch (res2)
                                    {
                                        case MessageBoxResult.Yes:
                                            EasyFilter_DeleteMail(mailItem);
                                            return;
                                        case MessageBoxResult.Cancel:
                                            MessageBox.Show("Canceled.");
                                            return;
                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    EasyFilter_DeleteMail(mailItem);
                                    return;
                                }
                                break;
                            case "moveto":
                                if (operationElement.GetAttribute("kakunin") != null)
                                {
                                    MessageBoxResult res2 = MessageBox.Show($"Would you like to move this mail?", "Confirmation", MessageBoxButton.YesNoCancel);
                                    switch (res2)
                                    {

                                        case MessageBoxResult.Yes:
                                            context?.HideMail(mailItem.EntryID, mailItem.Parent.StoreID);
                                            OutlookUtil.MoveMail(mailItem, operationElement.InnerText);
                                            return;
                                        case MessageBoxResult.Cancel:
                                            MessageBox.Show("Canceled.");
                                            return;
                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    context?.HideMail(mailItem.EntryID, mailItem.Parent.StoreID);
                                    OutlookUtil.MoveMail(mailItem, operationElement.Value);
                                    return;
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }
    }
}
