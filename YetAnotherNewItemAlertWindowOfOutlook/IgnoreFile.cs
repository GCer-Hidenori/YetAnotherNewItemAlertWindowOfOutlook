using System;
using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    internal class IgnoreFile
    {
        private static readonly string ignoreDirPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "excluded_ids");

        public static bool Exists(string entryID)
        {
            if (Directory.Exists(ignoreDirPath))
            {
                return File.Exists(Path.Combine(ignoreDirPath, entryID));
            }
            return false;
        }
        public static void Add(string entryID)
        {
            if (!Directory.Exists(ignoreDirPath))
            {
                Directory.CreateDirectory(ignoreDirPath);
            }
            File.Create(Path.Combine(ignoreDirPath, entryID));
        }
    }
}
