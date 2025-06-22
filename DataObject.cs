using System.Collections.Generic;

namespace CofyDev.Xml.Doc
{
    public partial class DataObject : Dictionary<string, object>
    {
        public DataObject() : base()
        {
        }

        public DataObject(int capacity) : base(capacity)
        {
        }
    }

    public class DataContainer : List<DataObject>
    {
    }
}