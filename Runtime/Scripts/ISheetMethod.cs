namespace HHG.GoogleSheets.Runtime
{
    public interface ISheetMethod
    {
        public string ColumnName { get; }
        public System.Type FieldType { get; }
    }
}