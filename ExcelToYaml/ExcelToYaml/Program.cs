using System;
using System.Diagnostics;
using System.IO;

namespace ExcelToYaml
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(".xlsx or .xls >>> .yaml 自動檔案轉換啟動..." + Environment.NewLine + 
                              "----------------------------------------------");

            (string loadExcelFolder, string saveYamlFolder) = GetEnvironmentVariables();

            Console.WriteLine("----------------------------------------------" + Environment.NewLine + 
                              "Aligala 遊戲資料自動轉換啟動..." + Environment.NewLine +
                              "----------------------------------------------");

            FileChangeEvent FileChangeEvent = new FileChangeEvent(loadExcelFolder, saveYamlFolder);
        }

        private static (string loadFolder, string saveFolder) GetEnvironmentVariables()
        {
            Console.Write("請輸入Excel檔案讀取目錄(若不輸入則讀取自本資料夾內): ");
            var getLoadFolder = Console.ReadLine();
            if (String.IsNullOrWhiteSpace(getLoadFolder))
            {
                getLoadFolder = Environment.CurrentDirectory;
            }
            else
            {
                while (true)
                {
                    if (!Directory.Exists(getLoadFolder))
                    {
                        Console.Write("所輸入的目錄不存在，請重新輸入: ");
                        getLoadFolder = Console.ReadLine();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            Console.Write("請輸入Yaml檔案存放目錄(若不輸入則儲存於本資料夾內): ");
            var getSaveFolder = Console.ReadLine();
            if (String.IsNullOrWhiteSpace(getSaveFolder))
            {
                getSaveFolder = Environment.CurrentDirectory;
            }
            else
            {
                while (true)
                {
                    if (!Directory.Exists(getSaveFolder))
                    {
                        Console.Write("所輸入的目錄不存在，請重新輸入: ");
                        getSaveFolder = Console.ReadLine();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return (loadFolder: getLoadFolder, saveFolder: getSaveFolder);
        }
    }
}
