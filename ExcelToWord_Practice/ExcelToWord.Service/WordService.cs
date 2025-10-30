using System;
using System.IO;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using ExcelToWord.Configuration;

namespace ExcelToWord.Service
{
    public class WordService : IWordService
    {
        private readonly Word.Application _wordApp;
        private readonly ExportSettings _settings;

        public WordService(ExportSettings settings)
        {
            _settings = settings;

            _wordApp = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
        }

        public Word.Document OpenOrCreate(string wordPath)
        {
            return File.Exists(wordPath)
                ? _wordApp.Documents.Open(wordPath)
                : _wordApp.Documents.Add();
        }

        public void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float imageWidthCm)
        {
            int maxRetries = 4;
            int currentRetry = 1;

            while (currentRetry < maxRetries)
            {
                try
                {
                    doc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);

                    if (_settings.InsertTitleBeforeImage)
                    {
                        var para = doc.Content.Paragraphs.Add(); // 避免使用 doc.Paragraphs.Add();
                        para.Range.Text = $"【{sheetName}】";
                        para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2); // 這句常常忘記
                        para.Range.InsertParagraphAfter(); // 還要記得換行
                    }

                    range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture); // 忘記要複製

                    doc.Activate();

                    _wordApp.Selection.EndKey(Unit: WdUnits.wdStory); // 整個忘記怎麼寫
                    _wordApp.Selection.Paste(); // 整個忘記怎麼寫wor

                    SetImageSize(doc, imageWidthCm); // 整個忘記怎麼寫

                    _wordApp.Selection.TypeParagraph(); // 整個忘記怎麼寫

                    break;
                }
                catch (Exception ex)
                {
                    currentRetry++;

                    if (currentRetry >= maxRetries)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Word圖片無法順利貼上，{ex.Message}，");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine($"第{currentRetry-1}次嘗試失敗，重新嘗試中：");
                        Thread.Sleep(300);
                    }
                }
            }
        }

        private void SetImageSize(Word.Document doc, float imageWidthCm) // 整個忘記怎麼寫
        {
            try
            {
                if (doc.InlineShapes.Count > 0)
                {
                    var shape = doc.InlineShapes[doc.InlineShapes.Count];
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Width = _wordApp.CentimetersToPoints(imageWidthCm);
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Word圖片大小無法變更，{ex.Message}，");
                Console.ResetColor();
            }
        }

        public void SaveAndClose(Word.Document doc, string wordPath)
        {
            try
            {
                doc.SaveAs2(wordPath);
                doc?.Close(false);

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Word發生異常，無法順利儲存關閉，{ex.Message}，");
                Console.ResetColor();
            }
            finally
            {
                if (doc != null)
                    try
                    {
                        Marshal.FinalReleaseComObject(doc);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Word物件釋放時發生警告：{ex.Message}");
                        Console.ResetColor();
                    }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void Close()
        {
            try
            {
                _wordApp?.Quit();
            }

            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Word執行檔發生異常，無法順利退出，{ex.Message}，");
                Console.ResetColor();
            }
            finally
            {

                if (_wordApp != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(_wordApp);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Word物件釋放時發生警告：{ex.Message}");
                        Console.ResetColor();
                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        } 
    }
}
