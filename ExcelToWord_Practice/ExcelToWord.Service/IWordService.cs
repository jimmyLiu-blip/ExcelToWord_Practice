using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord.Service
{
    public interface IWordService
    {
        Word.Document OpenOrCreate(string wordPath);

        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float imageWidthCm);

        void SaveAndClose(Word.Document doc, string wordPath);

        void Close();

        // 新增 Word 轉 PDF
        void ConvertWordToPdf(string wordPath);
    }
}