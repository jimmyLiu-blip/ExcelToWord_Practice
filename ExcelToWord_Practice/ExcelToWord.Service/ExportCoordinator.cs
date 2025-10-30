using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelToWord.Configuration;

namespace ExcelToWord.Service
{
    public class ExportCoordinator
    {
        private readonly ExportSettings _setting;
        private readonly IExcelService _excelService;
        private readonly IWordService _wordService;

        public ExportCoordinator(ExportSettings setting, IExcelService excelService, IWordService wordService)
        {
            _setting = setting;
            _excelService = excelService;
            _wordService = wordService;
        }

        public void Run()
        {
            Directory.CreateDirectory(_setting.OutputFolder);

            Excel.Workbook workbook = _excelService.Workbook;

            for (int i = _setting.StartSheetIndex; i <= workbook.Sheets.Count; i++) 
            {
                Excel.Worksheet ws = (Excel.Worksheet)workbook.Sheets[i]; // 這行很重要

                if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"略過隱藏工作表：{ws.Name}");
                    Console.ResetColor();
                    continue;
                }

                string sheetName = ws.Name;

                Console.WriteLine($"\n 正在處理工作表：{sheetName}");

                foreach (string rangeName in _setting.TargetNames)
                {

                    Excel.Range range = _excelService.GetRangeName(ws, rangeName);

                    if (range == null)
                    {
                        Console.WriteLine($"找不到命名範圍：{rangeName}(在{ws.Name})");
                        continue;
                    }

                    if (range.EntireRow.Hidden || range.EntireColumn.Hidden)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"略過隱藏範圍：{ws.Name}!{rangeName}");
                        Console.ResetColor();
                        continue;
                    }

                    // 先嘗試用「完整名稱」映射
                    string itemName;
                    if (_setting.PrefixToWordName.TryGetValue(rangeName, out var mappedFull))
                    {
                        itemName = mappedFull;
                    }
                    else
                    {
                        // 判斷是否為「尾段是數字」的樣式：例如 ACL_1、n77_10
                        string baseKey = rangeName;
                        int us = rangeName.LastIndexOf('_');
                        if (us > 0 && us < rangeName.Length - 1)
                        {
                            // 只有在最後一段是數字時，才把底線後面去掉當前綴
                            if (int.TryParse(rangeName.Substring(us + 1), out _))
                            {
                                baseKey = rangeName.Substring(0, us); // ACL_1 -> ACL, n77_10 -> n77
                            }
                        }

                        // 用前綴找對照表，找不到就用 baseKey 當檔名
                        itemName = _setting.PrefixToWordName.TryGetValue(baseKey, out var mappedPrefix)
                            ? mappedPrefix
                            : baseKey;
                    }

                    string wordPath = Path.Combine(_setting.OutputFolder, $"{itemName}.docx");

                    try
                    {
                        var doc = _wordService.OpenOrCreate(wordPath);
                        _wordService.InsertRangePicture(doc, sheetName, range, _setting.ImageWidthCm);
                        _wordService.SaveAndClose(doc, wordPath);
                        Console.WriteLine($"匯出成功：{rangeName} → {wordPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"❌ 匯出失敗：{rangeName}（在 {ws.Name}） - {ex.Message}");
                        Console.ResetColor();
                    }

                    Thread.Sleep(_setting.DelayMs);

                    Thread.Sleep(_setting.DelayMs);
                }
            }

            Console.WriteLine("\n 全部完成");

            _excelService.Close();
            _wordService.Close();

        }
    }
}