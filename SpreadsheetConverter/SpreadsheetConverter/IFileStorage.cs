using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetConverter.Files
{
    public interface IFileStorage
    {
        string StorageLocation { get; set; }

        void OpenFile();
    }
}
