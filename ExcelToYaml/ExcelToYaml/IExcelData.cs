using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToYaml
{
    interface IExcelData : IDisposable
    {
        IEnumerable<string> SheetNames { get; }
        SeedTableBase GetSeedTable(string sheetName, int columnNamesRowIndex = 2, int dataStartRowIndex = 3, IEnumerable<string> ignoreColumnNames = null, string keyColumnName = "ID", string versionColumnName = null);
        void Save();
        void SaveAs(string file);
    }
}
