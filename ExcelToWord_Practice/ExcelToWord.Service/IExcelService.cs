using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord.Service
{
    public interface IExcelService
    { 
        Excel.Workbook Workbook { get; }

        Excel.Range GetRangeName(Excel.Worksheet ws, string rangeName);

        void Close();
    }
}