using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Field)]
    public class SheetFieldAttribute : Attribute
    {
        public string ColumnName { get; }
        public string TransformMethod { get; }

        public SheetFieldAttribute(string columnName, string transformMethod = null)
        {
            ColumnName = columnName;
            TransformMethod = transformMethod;
        }
    }
}