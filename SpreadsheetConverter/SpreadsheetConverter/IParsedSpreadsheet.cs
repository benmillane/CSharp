using SpreadsheetConverter.Files;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetConverter.Interfaces
{
    public interface IParsedSpreadsheet<TEntity> where TEntity: IParsedRow
    {
        int Columns { get; set; }
        int Pages { get; set; }
        int InterfacePropertyCount { get; set; }
        Dictionary<string, int> Map { get; set; }
        List<string> ColumnHeaders { get; set; }
        List<TEntity> RowList { get; set; }
        List<TEntity> ParseSheet(IFileStorage storage);
        List<String> ObtainColumnHeaders(IFileStorage storage);
    }
}
