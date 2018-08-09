using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelToYaml
{
    public class SeedTableInterface
    {
        public class CannotContinueException : InvalidOperationException { }

        public static bool ExcelToYaml(string excelFilePath, string saveYamlFolder)
        {
            Console.WriteLine("開始檔案轉換");
            var startTime = DateTime.Now;

            CheckFileExists(excelFilePath);

            var excelFileInfo = EPPlus.ExcelData.FromFile(excelFilePath);

            var fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            // only use the first sheet 
            var firstSheetName = excelFileInfo.SheetNames.ToList()[0];
            var yamlName = fileName;

            var seedTable = GetSeedTable(excelFileInfo, firstSheetName);
            if (seedTable.Errors.Count != 0)
            {
                throw new CannotContinueException();
            }
            new YamlData(seedTable.ExcelToData(""), false, 0, 0, null, SeedYamlFormat.Hash, false, null)
                .WriteTo(yamlName, saveYamlFolder, ".yaml");

            var endTime = DateTime.Now;

            Console.WriteLine($"轉換成功 花費時間: {(endTime - startTime).TotalMilliseconds} ms" + Environment.NewLine);

            return true;
        }

        static SeedTableBase GetSeedTable(IExcelData excelData, string sheetName)
        {
            var seedTable = excelData.GetSeedTable(sheetName, 2, 3, null, "ID", null);
            if (seedTable.Errors.Count != 0)
            {
                var skipExceptions = seedTable.Errors.Where(error => error is NoIdColumnException);
                if (skipExceptions.Count() != 0)
                {
                    foreach (var error in skipExceptions)
                    {
                        Console.WriteLine($"      skip: {error.Message}");
                    }
                }
                else
                {
                    foreach (var error in seedTable.Errors)
                    {
                        Console.WriteLine($"      ERROR: {error.Message}");
                    }
                    throw new CannotContinueException();
                }
            }
            return seedTable;
        }

        private static ExcelPackage GetEPPlusExcelData(string file)
        {
            var fileInfo = new FileInfo(file);
            if (!fileInfo.Exists)
            {
                throw new FileNotFoundException(null, file);
            }
            return new ExcelPackage(fileInfo);
        }

        private static void CheckFileExists(string file)
        {
            if (!File.Exists(file))
            {
                Console.WriteLine($"file not found [{file}]");
                throw new CannotContinueException();
            }
        }
    }
}