using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Filter
    {
        private List<Condition> conditions = new();
        [XmlElement("Condition")]
        public List<Condition> Conditions { get => conditions; set => conditions = value; }

        public bool Evaluate(MailItem mailItem)
        {
            if (conditions.Count == 0)
            {
                return true;
            }
            else
            {
                return conditions.First().Evaluate(mailItem);
            }

        }
    }
}
