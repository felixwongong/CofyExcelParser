using System.Collections.Generic;

namespace CofyDev.Xml.Doc
{
    public class DataObject : Dictionary<string, object>
    {
        public DataObject subDataObject;

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