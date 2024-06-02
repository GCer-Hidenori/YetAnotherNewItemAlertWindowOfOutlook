using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Filter
    {
        private List<Condition> conditions = new();
        [XmlElement("Filter")]
        public List<Condition> Conditions { get => conditions; set => conditions = value; }

        public bool Evaluate(MailItem mailItem)
        {
            if(conditions.Count == 0)
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
