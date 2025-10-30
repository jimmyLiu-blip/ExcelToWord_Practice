using System;
using ExcelToWord.Configuration;
using ExcelToWord.Service;

namespace ExcelToWord_Practice
{
    public class Program
    {
        [STAThread] //遺漏沒寫到
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
                Console.WriteLine($"輸出範圍: {string.Join(",",settings.TargetNames)} ");  // 要使用Join讓裡面的字串隔開
                Console.WriteLine($"從第 {settings.StartSheetIndex} 張 sheet 開始匯出\n");

                IExcelService excelService = new ExcelService(settings.ExcelPath);
                IWordService wordService = new WordService(settings);

                ExportCoordinator coordinator = new ExportCoordinator(settings, excelService, wordService);

                coordinator.Run();

                Console.ForegroundColor = ConsoleColor.Green; // 忘記寫
                Console.WriteLine("\n所有作業已完成"); // 忘記寫
                Console.ResetColor();   // 忘記寫
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