using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelToWord.Service
{
    public class ExcelService : IExcelService
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _workbook;

        public ExcelService(string excelPath)
        {
            _excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
            };

            _workbook = _excelApp.Workbooks.Open(excelPath);
        }

        public Excel.Workbook Workbook => _workbook;

        public Excel.Range GetRangeName(Excel.Worksheet ws, string rangeName)
        { 
            Excel.Range range = null;

            try
            {
                range = ws.Names.Item(rangeName).RefersToRange;
            }
            catch
            {
                try
                {
                    var nameRange = _workbook.Names.Item(rangeName);
                    if (nameRange != null)
                    {
                        range = nameRange.RefersToRange;
                    }
                }
                catch
                { 
                    Console.ForegroundColor = ConsoleColor.Yellow;  // 顏色建議和錯誤分開
                    Console.WriteLine($"找不到命名範圍：{ws.Name}!{rangeName}"); // 輸出建議更直觀 
                    Console.ResetColor();
                }
            }

            return range;
        }

        public void Close()
        {
            try
            {
                _workbook?.Close(false);
                _excelApp?.Quit();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"關閉Excel檔案出現異常，{ex.Message}");
                Console.ResetColor();
            }
            finally
            {
                 if (_workbook != null)
                 {
                    try
                    {
                        Marshal.FinalReleaseComObject(_workbook);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Excel Workbook 物件釋放時發生警告，{ex.Message}");
                        Console.ResetColor();
                    }
                 }

                 if (_excelApp != null)
                 {  
                    try
                    {
                        Marshal.FinalReleaseComObject(_excelApp);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Excel Application 物件釋放時發生警告：{ex.Message}");
                        Console.ResetColor();
                    }
                }
                 
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}