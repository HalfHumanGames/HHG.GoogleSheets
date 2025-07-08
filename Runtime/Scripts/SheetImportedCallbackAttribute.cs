using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Method)]
    public class SheetImportedCallbackAttribute : Attribute
    {
        public string SpreadsheetId { get; }
        public string[] Gids { get; }

        public SheetImportedCallbackAttribute(string spreadsheetId, string gid = null)
        {
            SpreadsheetId = spreadsheetId;
            
            if (!string.IsNullOrEmpty(gid))
            {
                Gids = gid.Split(',', StringSplitOptions.RemoveEmptyEntries);
            }
        }
    }
}