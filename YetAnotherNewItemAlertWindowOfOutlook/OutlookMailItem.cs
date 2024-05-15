using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    class OutlookMailItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void RaisePropertyChanged([CallerMemberName] string? propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private string cc = "";
        private string categories = "";
        private string entry_id = "";
        private string flag_icon = "";
        private string recipient_names = "";
        private string recipient_addresses = "";
        private DateTime? reminder_time;
        private DateTime? receive_time;
        private string sender_email_address = "";
        private string sender_name = "";
        private DateTime? sent_on;
        private string subject = "";
        private string to = "";
        private string search_index = "";
        private Boolean unread;

        public string Cc 
        {
            get => cc;
            set
            {
                if(cc != value)
                {
                    cc = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public string Categories
        {
            get => categories;
            set
            {
                if(categories != value)
                {
                    categories = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public string EntryID {
            get => entry_id; 
            set
            {
                if(entry_id != value)
                {
                    entry_id = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string FlagIcon {
            get => flag_icon;
            set
            {
                if(flag_icon != value)
                {
                    flag_icon = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string RecipientNames {
            get => recipient_names; 
            set
            {
                if(recipient_names != value)
                {
                    recipient_names = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public string RecipientAddresses { 
            get => recipient_addresses;
            set
            {
                if(recipient_addresses != value)
                {
                    recipient_addresses = value;
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? ReminderTime {
            get => reminder_time;
            set
            {
                if(reminder_time != value)
                {
                    reminder_time = value;
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? ReceivedTime {
            get => receive_time;
            set
            {
                if(receive_time != value)
                {
                    receive_time = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string SenderEmailAddress {
            get => sender_email_address; 
            set
            {
                if(sender_email_address != value)
                {
                    sender_email_address = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string SenderName {
            get => sender_name;
            set
            {
                if(sender_name != value)
                {
                    sender_name = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? SentOn {
            get => sent_on;
            set
            {
                if(sent_on != value)
                {
                    sent_on = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string Subject 
        {
            get => subject;
            set
            {
                if(subject != value)
                {
                    subject = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }

        public string To {
            get => to;
            set
            {
                if(to != value)
                {
                    to = value;
                    RaisePropertyChanged();
                }
            }

        }
        public bool Unread {
            get => unread; 
            set
            {
                if(unread != value)
                {
                    unread = value;
                    RaisePropertyChanged();
                }
            }
        }
        private void RefreshSearchIndex()
        {
            search_index = cc.ToLower() + categories.ToLower() + recipient_names.ToLower() + sender_name.ToLower() + subject.ToLower();
        }
        public string SearchIndex 
        {
            get { 
                return search_index;
            }
            //set => search_index = value; 
        }

        public static OutlookMailItem CreateNew(MailItem mailItem)
        {
            var outlookmailitem = new OutlookMailItem()
            {
                cc = mailItem.CC,
                categories = mailItem.Categories,
                entry_id = mailItem.EntryID,
                flag_icon = (int)mailItem.FlagIcon switch
                {
                    0 => "No",
                    1 => "Purple",
                    2 => "Orange",
                    3 => "Green",
                    4 => "Yellow",
                    5 => "Blue",
                    6 => "Red",
                    _ => "",
                },
                subject = mailItem.Subject,
                to = mailItem.To,
                unread = mailItem.UnRead,
                sender_email_address = mailItem.SenderEmailAddress,
                sender_name = mailItem.SenderName
            };

         
            outlookmailitem.recipient_addresses = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Address));
            outlookmailitem.recipient_names = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Name));

            if (mailItem.ReminderTime > DateTime.Now.AddYears(100))
            {
                outlookmailitem.reminder_time = null;
            }
            else
            {
                outlookmailitem.reminder_time = mailItem.ReminderTime;
            }
            if (mailItem.ReceivedTime > DateTime.Now.AddYears(100))
            {
                outlookmailitem.receive_time = null;
            }
            else
            {
                outlookmailitem.receive_time = mailItem.ReceivedTime;
            }
            if (mailItem.SentOn > DateTime.Now.AddYears(100))
            {
                outlookmailitem.sent_on = null;
            }
            else
            {
                outlookmailitem.sent_on = mailItem.SentOn;
            }
            return outlookmailitem;

        }
        public static OutlookMailItem CreateNew(string entryID, Microsoft.Office.Interop.Outlook.Application outlook)
        {
            MailItem mailitem = OutlookUtil.GetMail(entryID, outlook);
            return CreateNew(mailitem);
        }

        public static void Reload(OutlookMailItem outlookmailitem, Microsoft.Office.Interop.Outlook.Application outlook)
        {
            MailItem mailitem = OutlookUtil.GetMail(outlookmailitem.EntryID, outlook);
            outlookmailitem.Categories = mailitem.Categories;
            outlookmailitem.FlagIcon = (int)mailitem.FlagIcon switch
            {
                0 => "No",
                1 => "Purple",
                2 => "Orange",
                3 => "Green",
                4 => "Yellow",
                5 => "Blue",
                6 => "Red",
                _ => "",
            };
            outlookmailitem.Unread = mailitem.UnRead;
            if (mailitem.ReminderTime > DateTime.Now.AddYears(100))
            {
                outlookmailitem.ReminderTime = null;
            }
            else
            {
                outlookmailitem.ReminderTime = mailitem.ReminderTime;
            }
            if (mailitem.ReceivedTime > DateTime.Now.AddYears(100))
            {
                outlookmailitem.ReceivedTime = null;
            }
            else
            {
                outlookmailitem.ReceivedTime = mailitem.ReceivedTime;
            }
            if (mailitem.SentOn > DateTime.Now.AddYears(100))
            {
                outlookmailitem.SentOn = null;
            }
            else
            {
                outlookmailitem.SentOn = mailitem.SentOn;
            }
        }
    }
}
