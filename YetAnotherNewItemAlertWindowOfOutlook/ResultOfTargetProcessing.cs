using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class ResultOfTargetProcessing
    {
        /*
        // duplicate / for update
        private List<string> list_duplication_entryid = new();
        // only target_processing / add to OutlookMailItemCollection
        */
        private List<string> list_new_entry_id = new();

        /*
        // only OutlookMailItemCollection / delete from OutlookMailItemCollection
        private List<string> list_deleted_entry_id = new();

        public List<string> List_duplication_entryid { get => list_duplication_entryid; set => list_duplication_entryid = value; }
        public List<string> List_new_entry_id { get => list_new_entry_id; set => list_new_entry_id = value; }
        public List<string> List_deleted_entry_id { get => list_deleted_entry_id; set => list_deleted_entry_id = value; }
        private bool newItemFound = false;

        public bool NewItemFound { get => newItemFound; set => newItemFound = value; }
        */
        private bool activateWindow = false;

        public bool ActivateWindow { get => activateWindow; set => activateWindow = value; }
        public List<string> List_new_entry_id { get => list_new_entry_id; set => list_new_entry_id = value; }
    }
}
