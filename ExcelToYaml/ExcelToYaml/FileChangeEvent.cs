using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToYaml
{
    class FileChangeEvent
    {
        private static string[] extensions = { ".xlsx", ".xls" };
        private string LoadExcelFolder, SaveYamlFolder;

        public FileChangeEvent(string loadExcelFolder, string saveYamlFolder)
        {
            LoadExcelFolder = loadExcelFolder;
            SaveYamlFolder = saveYamlFolder;

            var watcher = new FileSystemWatcher
            {
                Path = LoadExcelFolder,
                NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                               | NotifyFilters.FileName | NotifyFilters.DirectoryName,
                // watch excel files. Filter can only watch one category
                Filter = "*.*"
            };

            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Deleted += new FileSystemEventHandler(OnChanged);
            watcher.Renamed += new RenamedEventHandler(OnRenamed);

            watcher.EnableRaisingEvents = true;

            // wait - not to end
            new System.Threading.AutoResetEvent(false).WaitOne();
        }

        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            ExcelToYaml(e.Name , e.FullPath);
        }

        // Define the event handlers.
        private void OnRenamed(object source, RenamedEventArgs e)
        {
            ExcelToYaml(e.Name, e.FullPath);
        }

        private void ExcelToYaml(string name, string fullPath)
        {
            try
            {
                // The ~ file is when excel file is modifying
                if (Path.GetFileName(fullPath).StartsWith("~"))
                {
                    return;
                }

                // Added null check to avoid possible NullReferenceException.
                var ext = (Path.GetExtension(fullPath) ?? string.Empty).ToLower();

                if (extensions.Any(ext.Equals))
                {
                    // Specify what is done when a file is changed, created, or deleted.
                    Console.WriteLine($"偵測到Excel類型檔案變更: {name}");

                    SeedTableInterface.ExcelToYaml(fullPath, SaveYamlFolder);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"錯誤 錯誤訊息: {ex.ToString()}" + Environment.NewLine);
            }
        }
    }
}
