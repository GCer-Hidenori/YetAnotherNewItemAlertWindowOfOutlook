using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    class OutlookMailItem : INotifyPropertyChanged
    {
        public OutlookMailItem()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }
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
        private string store_id = "";
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
        private string conversation_id = "";

        public void RefreshSearchIndex()
        {
            search_index = Strings.StrConv(cc + categories + recipient_names + sender_name + subject, VbStrConv.Wide, 0) ?? "";
        }

        public string Cc
        {
            get => cc;
            set
            {
                if (cc != value)
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
                if (categories != value)
                {
                    categories = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public string StoreID
        {
            get => store_id;
            set
            {
                if (store_id != value)
                {
                    store_id = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string EntryID
        {
            get => entry_id;
            set
            {
                if (entry_id != value)
                {
                    entry_id = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string FlagIcon
        {
            get => flag_icon;
            set
            {
                if (flag_icon != value)
                {
                    flag_icon = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string RecipientNames
        {
            get => recipient_names;
            set
            {
                if (recipient_names != value)
                {
                    recipient_names = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public string RecipientAddresses
        {
            get => recipient_addresses;
            set
            {
                if (recipient_addresses != value)
                {
                    recipient_addresses = value;
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? ReminderTime
        {
            get => reminder_time;
            set
            {
                if (reminder_time != value)
                {
                    reminder_time = value;
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? ReceivedTime
        {
            get => receive_time;
            set
            {
                if (receive_time != value)
                {
                    receive_time = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string SenderEmailAddress
        {
            get => sender_email_address;
            set
            {
                if (sender_email_address != value)
                {
                    sender_email_address = value;
                    RaisePropertyChanged();
                }
            }
        }
        public string SenderName
        {
            get => sender_name;
            set
            {
                if (sender_name != value)
                {
                    sender_name = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }
        public DateTime? SentOn
        {
            get => sent_on;
            set
            {
                if (sent_on != value)
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
                if (subject != value)
                {
                    subject = value;
                    RefreshSearchIndex();
                    RaisePropertyChanged();
                }
            }
        }

        public string To
        {
            get => to;
            set
            {
                if (to != value)
                {
                    to = value;
                    RaisePropertyChanged();
                }
            }

        }
        public bool Unread
        {
            get => unread;
            set
            {
                if (unread != value)
                {
                    unread = value;
                    RaisePropertyChanged();
                }
            }
        }

        public string SearchIndex
        {
            get
            {
                return search_index;
                //return cc + categories + recipient_names + sender_name + subject;
            }
            //set => search_index = value; 
        }

        public string ConversationId { get => conversation_id; set => conversation_id = value; }

        public static OutlookMailItem CreateNew(MailItem mailItem, string storeID)
        {
            var outlookMailItem = new OutlookMailItem()
            {
                cc = mailItem.CC,
                categories = mailItem.Categories,
                //store_id = mailItem.Parent.StoreID, //here
                store_id = storeID, //here
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
                sender_name = mailItem.SenderName,
                ConversationId = mailItem.ConversationID
            };
            outlookMailItem.RefreshSearchIndex();


            outlookMailItem.recipient_addresses = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Address));
            outlookMailItem.recipient_names = String.Join(";", mailItem.Recipients.Cast<Recipient>().ToList().Select(recipient => recipient.Name));

            if (mailItem.ReminderTime > DateTime.Now.AddYears(100))
            {
                outlookMailItem.reminder_time = null;
            }
            else
            {
                outlookMailItem.reminder_time = mailItem.ReminderTime;
            }
            if (mailItem.ReceivedTime > DateTime.Now.AddYears(100))
            {
                outlookMailItem.receive_time = null;
            }
            else
            {
                outlookMailItem.receive_time = mailItem.ReceivedTime;
            }
            if (mailItem.SentOn > DateTime.Now.AddYears(100))
            {
                outlookMailItem.sent_on = null;
            }
            else
            {
                outlookMailItem.sent_on = mailItem.SentOn;
            }
            return outlookMailItem;

        }
        public static OutlookMailItem CreateNew(string storeID, string entryID, Microsoft.Office.Interop.Outlook.Application outlook)
        {
            MailItem mailitem = OutlookUtil.GetMail(storeID, entryID, outlook);
            return CreateNew(mailitem, storeID);
        }

        public static void Reload(OutlookMailItem outlookMailItem, Microsoft.Office.Interop.Outlook.Application outlook)
        {
            MailItem mailitem = OutlookUtil.GetMail(outlookMailItem.StoreID, outlookMailItem.EntryID, outlook);
            outlookMailItem.Categories = mailitem.Categories;
            outlookMailItem.FlagIcon = (int)mailitem.FlagIcon switch
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
            outlookMailItem.Unread = mailitem.UnRead;
            if (mailitem.ReminderTime > DateTime.Now.AddYears(100))
            {
                outlookMailItem.ReminderTime = null;
            }
            else
            {
                outlookMailItem.ReminderTime = mailitem.ReminderTime;
            }
            if (mailitem.ReceivedTime > DateTime.Now.AddYears(100))
            {
                outlookMailItem.ReceivedTime = null;
            }
            else
            {
                outlookMailItem.ReceivedTime = mailitem.ReceivedTime;
            }
            if (mailitem.SentOn > DateTime.Now.AddYears(100))
            {
                outlookMailItem.SentOn = null;
            }
            else
            {
                outlookMailItem.SentOn = mailitem.SentOn;
            }
        }
    }
}
