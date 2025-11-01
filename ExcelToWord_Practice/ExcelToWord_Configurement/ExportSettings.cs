using System;
using System.Collections.Generic;

namespace ExcelToWord.Configuration
{
    public class ExportSettings
    {
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        public string[] TargetNames { get; set; } = { "ACL_1", "ACLN_1" };

        public int StartIndexSheet { get; set; } = 4;

        public float ImageWidthCm { get; set; } = 18;

        public int DelayMs { get; set; } = 150;

        public bool InsertTitleBeforeImage { get; set; } = false;

        // 如果要自定義Word名稱，建議使用Dictionary的key、value兩兩一組
        public Dictionary<string,string> PrefixToWordName { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "ACL", "ACP" },
            { "ACLN", "Leakage_Power_No_Carrier" },
            { "FTol", "Frequency_Tolerance"},
            { "Inter", "Intermodulation" },
            { "OBE", "Out_Band_Emission"},
            { "OCB", "OBW"},
            { "Power","Antenna_Power"},
            { "RX", "Secondary_Emission"},
            { "Spurious", "Spurious_Emission"},
            { "n77", "RF_Specification"},
            { "n40", "RF_Specification"},
            { "Test_Condition", "Test_Condition"},
            { "Voltage_Variation", "Voltage_Variation"},
            { "Result", "Test_Summary"}
        };
    }
}


