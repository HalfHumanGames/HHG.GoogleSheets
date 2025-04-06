using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetAttribute : Attribute
    {
        public string SpreadsheetId { get; }
        public string Gid { get; }

        public SheetAttribute(string spreadsheetId, string gid = null)
        {
            SpreadsheetId = spreadsheetId;
            Gid = gid;
        }
    }
}