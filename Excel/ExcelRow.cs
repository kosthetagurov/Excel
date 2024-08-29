namespace Excel
{
    public class ExcelRow
    {
        public Dictionary<string, string> Cells { get; set; } = new Dictionary<string, string>();

        public string this[string columnName]
        {
            get => Cells.ContainsKey(columnName) ? Cells[columnName] : null;
            set => Cells[columnName] = value;
        }

        public bool IsRowEmpty()
        {
            return Cells.Values.All(string.IsNullOrEmpty);
        }
    }
}
