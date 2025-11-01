using ExcelToWord.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToWord.Service
{
    public class ExportCoordinator
    { 
        private readonly ExportSettings _settings;
        private readonly IExcelService _excelService;
        private readonly IWordService _wordService;

        // HashSet<string> 是一種不允許重複的集合
        // HashSet.Contains 幾乎瞬間完成；List.Contains 要一個個比對，資料大時會慢。
        private readonly HashSet<string> _initializedWordFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public ExportCoordinator(ExportSettings settings, IExcelService excelService, IWordService wordService)
        { 
            _settings = settings;
            _excelService = excelService;
            _wordService = wordService;
        }

        public void Run()
        {
            Directory.CreateDirectory(_settings.OutputFolder);

            Excel.Workbook workbook = _excelService.Workbook;

            for (int i = _settings.StartIndexSheet; i <= workbook.Sheets.Count; i++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)workbook.Sheets[i];  // 這行很重要，但是一直忘記

                if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)  // 補上隱藏工作表單不轉出
                { 
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"略過隱藏工作表:{ws.Name}");
                    Console.ResetColor();
                    continue;
                }

                string sheetName = ws.Name;

                Console.WriteLine($"\n 正在處理工作表單:{sheetName}");

                foreach (string rangeName in _settings.TargetNames) // 忘記怎麼寫
                {
                    Excel.Range range = _excelService.GetRangeName(ws, rangeName); // 忘記怎麼寫

                    if (range == null) // 遺漏判斷range是否為空
                    {
                        Console.WriteLine($"找不到命名範圍：{rangeName}在{ws.Name}");
                        continue;
                    }

                    if (range.EntireColumn.Hidden || range.EntireRow.Hidden) // 新增表格隱藏時，不轉出
                    { 
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"略過隱藏範圍:{ws.Name}!{rangeName}");
                        Console.ResetColor();
                        continue;
                    }

                    string itemName; //新增如何判斷檔名要怎麼取捨

                    if (_settings.PrefixToWordName.TryGetValue(rangeName, out var mappedFull)) // 使用鏡射Key、Value
                    {
                        itemName = mappedFull;
                    }
                    else
                    {
                        string baseKey = rangeName;
                        int us = rangeName.LastIndexOf("_");
                        if ( us > 0 && us < rangeName.Length -1)
                        {
                            // Substring取出_之後的數字， _是丟棄變數，不關心實際轉出來的數字，只要知道能不能轉成功
                            if (int.TryParse(rangeName.Substring(us + 1), out _))
                            {
                                baseKey = rangeName.Substring(0, us);
                            }
                        }

                        itemName = _settings.PrefixToWordName.TryGetValue(baseKey, out var mappedPrefix)
                            ? mappedPrefix
                            : baseKey;
                    }

                    string wordPath = Path.Combine(_settings.OutputFolder, $"{itemName}.docx");

                    try
                    {
                        if (!_initializedWordFiles.Contains(wordPath))
                        {
                            if (File.Exists(wordPath))
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"偵測到舊檔，將刪除覆蓋:{wordPath}");
                                Console.ResetColor();
                            }
                            _initializedWordFiles.Add(wordPath); 
                        }

                        Word.Document doc = _wordService.OpenOrCreate(wordPath);

                        _wordService.InsertRangePicture(doc, sheetName, range, _settings.ImageWidthCm);

                        _wordService.SaveAndClose(doc, wordPath);

                        Console.WriteLine($"匯出成功: {rangeName} → {wordPath} ");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"匯出失敗：{rangeName}（在 {ws.Name}） - {ex.Message}");
                        Console.ResetColor();
                    }

                    Thread.Sleep(_settings.DelayMs); 
                } 
            }

            Console.WriteLine("\n全部 Word 檔匯出完成，開始轉換 PDF...");

            foreach (var wordFile in _initializedWordFiles)
            {
                _wordService.ConvertWordToPdf(wordFile);
            }

            Console.WriteLine("\n 全部完成");

            _excelService.Close();
            _wordService.Close();
        }
    }
}