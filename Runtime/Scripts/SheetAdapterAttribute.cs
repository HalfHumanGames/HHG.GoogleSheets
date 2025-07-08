using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    public class SheetAdapterAttribute : Attribute, ISheetMethod
    {
        public string ColumnName { get; }
        public Type FieldType { get; }

        public SheetAdapterAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        public SheetAdapterAttribute(Type fieldType)
        {
            FieldType = fieldType;
        }
    }
}