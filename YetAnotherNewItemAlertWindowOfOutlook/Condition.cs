using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public enum ConditionType
    {
        And,
        Or,
        Not,
        Subject,
        Body,
        To,
        Cc,
        SenderAddress,
        SenderName,
        RecipientNames,
        RECIPIENTADDRESSES,
        ATTACHMENT
    }

    public class Condition
    {
        private ConditionType type;
        private List<Condition> conditions = new();
        private string? value;

        [XmlAttribute("type")]
        public ConditionType Type { get => type; set => type = value; }

        [XmlElement("Condition")]
        public List<Condition> Conditions { get => conditions; set => conditions = value; }

        [XmlAttribute("value")]
        public string? Value { get => value; set => this.value = value; }

        public bool Evaluate(MailItem mailItem)
        {
            switch (type)
            {
                case ConditionType.And:
                    return conditions.All(c => c.Evaluate(mailItem));
                case ConditionType.Or:
                    return conditions.Any(c => c.Evaluate(mailItem));
                case ConditionType.Not:
                    return !conditions[0].Evaluate(mailItem);
                case ConditionType.Subject:
                    return (mailItem.Subject ?? "").Contains(Value ?? "");
                case ConditionType.Body:
                    return (mailItem.Body ?? "").Contains(Value ?? "");
                case ConditionType.To:
                    return (mailItem.To ?? "").Contains(Value ?? "");
                case ConditionType.Cc:
                    return (mailItem.CC ?? "").Contains(Value ?? "");
                case ConditionType.SenderAddress:
                    return (mailItem.SenderEmailAddress ?? "").Contains(Value ?? "");
                case ConditionType.SenderName:
                    return (mailItem.SenderName ?? "").Contains(Value ?? "");
                case ConditionType.RecipientNames:
                    return mailItem.Recipients.Cast<Recipient>().Any(r => r.Name.Contains(Value ?? ""));
                case ConditionType.RECIPIENTADDRESSES:
                    return mailItem.Recipients.Cast<Recipient>().Any(r => r.Address.Contains(Value ?? ""));
                case ConditionType.ATTACHMENT:
                    return mailItem.Attachments.Cast<Attachment>().Any(a => a.FileName.Contains(Value ?? ""));
                default:
                    return false;
            }
        }
    }

}
