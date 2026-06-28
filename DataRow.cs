using System.Collections.Generic;

namespace CofyDev.Xml.Doc
{
    public partial class DataRow : Dictionary<string, object>
    {
        public DataRow() : base()
        {
        }

        public DataRow(int capacity) : base(capacity)
        {
        }
    }

    public class DataTable : List<DataRow>
    {
    }
}