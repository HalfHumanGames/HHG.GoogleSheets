using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetAttribute : Attribute
    {
        public string SpreadsheetId { get; }
        public string Gid { get; }
        public NameCase Casing { get; }

        public SheetAttribute(string spreadsheetId, string gid = null, NameCase casing = NameCase.None)
        {
            SpreadsheetId = spreadsheetId;
            Gid = gid;
            Casing = casing;
        }
    }
}