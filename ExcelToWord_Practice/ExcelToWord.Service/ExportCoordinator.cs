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

            for(int i = _setting.StartSheetIndex; i <= workbook.Sheets.Count; i++) 
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

                    string itemName = rangeName.Contains("_")
                        ? rangeName.Split('_')[0] 
                        : rangeName;

                    string wordPath = Path.Combine( _setting.OutputFolder, $"{itemName}.docx");

                    var doc = _wordService.OpenOrCreate(wordPath);

                    _wordService.InsertRangePicture(doc, sheetName, range, _setting.ImageWidthCm);

                    _wordService.SaveAndClose(doc, wordPath);

                    Console.WriteLine($"匯出 {rangeName} => {wordPath}");
                    
                    Thread.Sleep(_setting.DelayMs);
                }
            }

            Console.WriteLine("\n 全部完成");

            _excelService.Close();
            _wordService.Close();

        }
    }
}