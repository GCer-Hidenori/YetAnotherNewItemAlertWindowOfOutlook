using System;
using System.Collections.Generic;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public enum ConditionType
    {
        AND,
        OR,
        NOT,
        SUBJECT,
        BODY,
        TO,
        CC,
        SENDEREMAILADDRESS,
        SENDERNAME,
        RECIPIENTNAMES,
        RECIPIENTADDRESSES,
        ATTACHMENT
    }

    public class Condition
    {
        private ConditionType type;
        private List<Condition> conditions;
        private string value;

        [XmlAttribute("type")]
        public ConditionType Type { get => type; set => type = value; }

        [XmlElement("Condition")]
        public List<Condition> Conditions { get => conditions; set => conditions = value; }

        [XmlAttribute("value")]
        public string Value { get => value; set => this.value = value; }

        public bool Evaluate(MailItem mailItem)
        {
            switch (type)
            {
                case ConditionType.AND:
                    return conditions.All(c => c.Evaluate(mailItem));
                case ConditionType.OR:
                    return conditions.Any(c => c.Evaluate(mailItem));
                case ConditionType.NOT:
                    return !conditions[0].Evaluate(mailItem);
                case ConditionType.SUBJECT:
                    return mailItem.Subject.Contains(Value);
                case ConditionType.BODY:
                    return mailItem.Body.Contains(Value);
                case ConditionType.TO:
                    return mailItem.To.Contains(Value);
                case ConditionType.CC:
                    return mailItem.CC.Contains(Value);
                case ConditionType.SENDEREMAILADDRESS:
                    return mailItem.SenderEmailAddress.Contains(Value);
                case ConditionType.SENDERNAME:
                    return mailItem.SenderName.Contains(Value);
                case ConditionType.RECIPIENTNAMES:
                    return mailItem.Recipients.Cast<Recipient>().Any(r => r.Name.Contains(Value));
                case ConditionType.RECIPIENTADDRESSES:
                    return mailItem.Recipients.Cast<Recipient>().Any(r => r.Address.Contains(Value));
                case ConditionType.ATTACHMENT:
                    return mailItem.Attachments.Cast<Attachment>().Any(a => a.FileName.Contains(Value));
                default:
                    return false;
            }
        }
    }

}
