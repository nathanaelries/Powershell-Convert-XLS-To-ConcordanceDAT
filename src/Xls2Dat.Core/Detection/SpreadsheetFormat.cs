namespace Xls2Dat.Core.Detection
{
    public enum SpreadsheetFormat
    {
        Unknown,
        Csv,        // .csv / .tsv / .txt — delimited text
        OpenXml,    // .xlsx / .xlsm — modern Excel (zip + OpenXML)
        LegacyXls,  // .xls — pre-2007 BIFF
        OpenDocument, // .ods — LibreOffice/OpenOffice
        AppleNumbers  // .numbers — Apple iWork
    }
}
