using System;

namespace HHG.GoogleSheets.Runtime
{
    [AttributeUsage(AttributeTargets.Field)]
    public class SheetFieldAttribute : Attribute
    {
        public string ColumnName { get; }
        public string TransformMethod { get; }
        public string AdapterMethod { get; }

        public SheetFieldAttribute(string columnName = null, string transformMethod = null, string adapterMethod = null)
        {
            ColumnName = columnName;
            TransformMethod = transformMethod;
            AdapterMethod = adapterMethod;
        }
    }
}