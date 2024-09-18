using System.Reflection;

namespace CofyDev.Xml.Doc;

public static class DataObjectExtension
{
    public static T DecodeAs<T>(this CofyXmlDocParser.DataObject dataObject, Action<FieldInfo, object, string> propertyDecodeSetter)
    {
        var objType = typeof(T);
        
        var obj = Activator.CreateInstance<T>();
        if (obj == null)
        {
            throw new ArgumentException($"Cannot create instance type ({typeof(T)})");
        }

        var fields = objType.GetFields();

        foreach (var (key, value) in dataObject)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException(nameof(key), "dataObject key is empty");
            }
            
            var field = objType.GetField(key, BindingFlags.Instance | BindingFlags.Public);
            if (field == null)
            {
                throw new KeyNotFoundException($"{objType} public field not found for key ({key})");
            }

            propertyDecodeSetter(field, obj, value);
        }

        return obj;
    }

    public static void SetDecodePropertyValue(FieldInfo fieldInfo, object obj, string rawValue)
    {
        var propertyType = fieldInfo.FieldType;

        bool parsable = false;
        if (propertyType == typeof(bool))
        {
            parsable = bool.TryParse(rawValue, out var value);
            if (parsable) fieldInfo.SetValue(obj, value);
        }
        else if (propertyType == typeof(int))
        {
            parsable = int.TryParse(rawValue, out var value);
            if (parsable) fieldInfo.SetValue(obj, value);
        }
        else if (propertyType == typeof(float))
        {
            parsable = float.TryParse(rawValue, out var value);
            if (parsable) fieldInfo.SetValue(obj, value);
        } 
        else if (propertyType == typeof(double))
        {
            parsable = double.TryParse(rawValue, out var value);
            if (parsable) fieldInfo.SetValue(obj, value);
        } 
        else if (propertyType.IsEnum)
        {
            parsable = Enum.TryParse(propertyType, rawValue, out var value);
            if(parsable) fieldInfo.SetValue(obj, value);
        }
        else
        {
            parsable = true;
            fieldInfo.SetValue(obj, rawValue);
        }

        if (!parsable)
        {
            throw new ArgumentException(
                $"dataObject value ({rawValue}) cannot parse to {obj.GetType()}'s {propertyType} field {fieldInfo.Name}");
        }
    }
}