using System.Collections.Generic;
using System.Xml.Serialization;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Rule
    {
        private Condition condition = new();
        public Condition? Condition { get => condition; set => condition = value; }

        [XmlArray("Actions")]
        private List<Action> actions = new();
        public List<Action> Actions { get => actions; set => actions = value; }

    }
}
