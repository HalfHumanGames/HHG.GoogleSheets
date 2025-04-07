using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    public class SheetTransformAttribute : Attribute
    {
        public string ColumnName { get; }

        public SheetTransformAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    } 
}