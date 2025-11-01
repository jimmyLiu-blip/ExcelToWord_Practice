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

        public WordService(ExportSettings setting)
        {
            _settings = setting;

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
            int MaxRetries = 3;
            int CurrentRetry = 0;

            while (CurrentRetry < MaxRetries)
            {
                try
                {
                    doc.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // 控制是否要新增標題在圖片前面
                    if (_settings.InsertTitleBeforeImage)
                    {
                        var para = doc.Content.Paragraphs.Add();
                        para.Range.Text = $"【{sheetName}】";
                        para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2); // 遺漏標題樣式
                        para.Range.InsertParagraphAfter();
                    }

                    range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                    doc.Activate();

                    _wordApp.Selection.EndKey(Unit: WdUnits.wdStory);  // 錯誤，之前寫的是變更標題
                    _wordApp.Selection.Paste();
                    SetImageSize(doc, imageWidthCm); // 順序不可變更，會導致BUG產生
                    _wordApp.Selection.TypeParagraph();   

                    break;
                }
                catch (Exception ex)
                {
                    CurrentRetry++;

                    if (CurrentRetry >= MaxRetries)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"發生異常，無法在Word中插入圖片，{ex.Message}");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine($"目前重新匯出第{CurrentRetry + 1}次，重新嘗試中");
                        Thread.Sleep(300);
                    }
                }
            }
        }

        private void SetImageSize(Word.Document doc, float imageWidthCm)
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
                Console.WriteLine($"發生異常，無法變更圖片大小，{ex.Message}");
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
                Console.WriteLine($"關閉Word檔案出現異常，{ex.Message}");
                Console.ResetColor();
            }
            finally
            {
                if (doc != null)
                {
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
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void Close()
        {
            try
            {
                _wordApp?.Quit();  // NULL允許
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"關閉Word執行檔案出現異常，{ex.Message}");
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

        // 新增 Word檔案轉 PDF檔案
        public void ConvertWordToPdf(string wordPath)
        {
            if (!File.Exists(wordPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"找不到 Word 檔案，{wordPath}");
                Console.ResetColor();
            }

            // 把字串 wordPath 代表的檔案路徑，把副檔名改成 .pdf，並回傳改好的新路徑字串，存到 pdfPath
            // 是純字串處理，不會真的改動硬碟上的檔案。
            string pdfPath = Path.ChangeExtension(wordPath, "pdf");

            Word.Application app = null;
            Word.Document doc = null;
            try
            {
                app = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = WdAlertLevel.wdAlertsNone,
                };

                doc = app.Documents.Open(wordPath, ReadOnly: true, Visible: false);

                doc.ExportAsFixedFormat(
                    pdfPath, // 要輸出的檔案完整路徑
                    Word.WdExportFormat.wdExportFormatPDF, // 匯出格式：wdExportFormatPDF 或 wdExportFormatXPS
                    OpenAfterExport: false, // 匯出後是否自動開啟 PDF
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint, // 轉檔品質：印刷 vs 螢幕 wdExportOptimizeForPrint → 高品質印刷用（字型完整）； wdExportOptimizeForOnScreen → 檔案小但品質略低
                    Range: Word.WdExportRange.wdExportAllDocument, // 要匯出的範圍，wdExportAllDocument（整份文件）或 wdExportFromTo（指定頁)
                    From: 0,
                    To: 0,
                    Item: Word.WdExportItem.wdExportDocumentContent, // 匯出內容類型，wdExportDocumentContent（僅文件內容）；wdExportDocumentWithMarkup（包含修訂）
                    IncludeDocProps: true, // 是否包含文件屬性（作者、標題）
                    KeepIRM: true, // 是否保留資訊權限管理（IRM）
                    CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, // 是否根據 Word 標題自動產生書籤，wdExportCreateHeadingBookmarks → 根據標題建立書籤；wdExportCreateNoBookmarks → 不建立書籤
                    DocStructureTags: true, // 是否建立 PDF 結構標籤（方便閱讀器理解章節）
                    BitmapMissingFonts: true, // 若缺字型，是否以點陣圖嵌入 
                    UseISO19005_1: false //是否轉成 PDF/A-1b 標準格式（長期保存）
                );

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"成功轉換為PDF:{pdfPath}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($" Word 轉 PDF 失敗，{ex.Message}");
                Console.ResetColor();
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                    Marshal.FinalReleaseComObject(doc);
                }

                if (app != null)
                { 
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}