using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public enum ErrorType
    {
        SettingFileNotFound,
        SettingFileLoadError,
        InvalidFilterElementName,
        InvalidTargetFolderPath,
        StoreNotFound,
        NoFolderFoundError,
        SampleSettingFileLoadError,
        ActionCreateFileError
    }

    public class YError : Exception
    {
        public ErrorType ErrorType { get; set; }
        public YError()
            : base()
        {
        }

        public YError(string message)
            : base(message)
        {
        }
        public YError(ErrorType errorType):base(errorType.ToString())
        {
            this.ErrorType = errorType;
        }
        public YError(ErrorType errorType,string message):base(errorType.ToString()+message)
        {
            this.ErrorType = errorType;
        }
        
    }
}
