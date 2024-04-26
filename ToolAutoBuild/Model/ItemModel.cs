namespace ToolAutoBuild.Model
{
    public class ItemModel
    {
        public string Key { get; set; }

        public string Value { get; set; }

        public string Comment { get; set; }

        public ItemModel(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public ItemModel(string key, string value, string comment)
        {
            Key = key;
            Value = value;
            Comment = comment;
        }
    }
}
