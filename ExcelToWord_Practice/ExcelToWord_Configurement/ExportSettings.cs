
using System;
using System.Collections.Generic;

namespace ExcelToWord.Configuration
{
    public class ExportSettings
    {
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        public string[] TargetNames { get; set; } = { 
            "n77_10","n77_15","Test_Condition","Voltage_Variation","Result",
            "ACL_1", "ACL_2","ACL_3","ACL_4",
            "ACLN_1","ACLN_2","ACLN_3","ACLN_4",
            "FTol_1","FTol_2","FTol_3","FTol_4",
            "Inter_1","Inter_2","Inter_3","Inter_4",
            "OBE_1","OBE_2","OBE_3","OBE_4",
            "OCB_1","OCB_2","OCB_3","OCB_4",
            "Power_1","Power_2","Power_3","Power_4",
            "RX_1","RX_2","RX_3","RX_4",
            "Spurious_1","Spurious_2","Spurious_3","Spurious_4" 
        };

        public int StartSheetIndex { get; set; } = 4;

        public float ImageWidthCm { get; set; } = 18;

        public int DelayMs { get; set; } = 150;

        public bool InsertTitleBeforeImage { get; set; } = false;

        public Dictionary<string, string> PrefixToWordName { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
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

