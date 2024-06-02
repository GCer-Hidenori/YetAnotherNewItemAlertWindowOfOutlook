using System.Collections.Generic;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class ResultOfTargetProcessing
    {
        /*
        // duplicate / for update
        private List<string> list_duplication_entryid = new();
        // only target_processing / add to OutlookMailItemCollection
        */
        private List<MailID> list_new_entry_id = new();

        public List<MailID> List_new_mail_id { get => list_new_entry_id; set => list_new_entry_id = value; }
    }
}
