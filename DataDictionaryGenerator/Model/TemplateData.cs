using System.Collections.Generic;

namespace DataDictionaryGenerator.Model
{
    public class TemplateData
    {
        public string TableName { get; set; }
        public List<TableColumnInfo> TableColumnInfos { get; set; }
    }


    public class TableColumnInfo
    {
        public int Index { get; set; }
        public string TableName { get; set; }
        public string Name { get; set; }
        public string IsPrimary { get; set; }
        public string Type { get; set; }
        public string AllowNull { get; set; }
        public string Description { get; set; }
    }
}