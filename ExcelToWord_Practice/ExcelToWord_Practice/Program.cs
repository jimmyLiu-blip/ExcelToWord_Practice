using System;
using System.Runtime.InteropServices;
using ExcelToWord.Configuration;
using ExcelToWord.Service;

namespace ExcelToWord_Practice
{
    public class Program
    {
        [STAThread] 
        static void Main()
        {
            try
            {
                Console.WriteLine("===================");
                Console.WriteLine("開始匯出Excel至Word");
                Console.WriteLine("===================\n");

                ExportSettings settings = new ExportSettings();

                Console.WriteLine($"Excel 路徑為: {settings.ExcelPath} ");
                Console.WriteLine($"Word 輸出資料夾路徑為: {settings.OutputFolder} ");
                Console.WriteLine($"輸出範圍為：{string.Join(",", settings.TargetNames)}");
                Console.WriteLine($"從第 {settings.StartIndexSheet} 張 sheet 開始匯出\n");

                IExcelService excelService = new ExcelService(settings.ExcelPath);
                IWordService wordService = new WordService(settings);

                ExportCoordinator coordinator = new ExportCoordinator(settings, excelService, wordService);

                coordinator.Run();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n所有作業已完成"); 
                Console.ResetColor();  
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"發生異常狀況，{ex.Message}");
                Console.WriteLine($"異常狀況路徑，\n{ex.StackTrace}");
                Console.ResetColor();
            }
            Console.WriteLine("\n按任意鍵離開");
            Console.ReadKey();
        }
    }
}