using System;
using System.Collections.Generic;
using System.Reflection;

namespace CofyDev.Xml.Doc
{
    public class DataObjectEncoder: IDisposable
    {
        private Dictionary<string, FieldInfo> _fieldCache = new();

        public virtual CofyXmlDocParser.DataObject Encode(object obj)
        {
            throw new NotImplementedException();
        }

        public virtual T DecodeAs<T>(CofyXmlDocParser.DataObject dataObject,
            Action<FieldInfo, object, string> propertyDecodeSetter)
        {
            var objType = typeof(T);

            var obj = Activator.CreateInstance<T>();
            if (obj == null)
            {
                throw new ArgumentException($"Cannot create instance type ({typeof(T)})");
            }

            foreach (var (key, value) in dataObject)
            {
                if (string.IsNullOrEmpty(key))
                {
                    throw new ArgumentNullException(nameof(key), "dataObject key is empty");
                }

                if (!_fieldCache.TryGetValue(key, out var field))
                {
                    field = objType.GetField(key, BindingFlags.Instance | BindingFlags.Public);
                    _fieldCache[key] = field;
                }

                if (field == null)
                {
                    throw new KeyNotFoundException($"{objType} public field not found for key ({key})");
                }

                propertyDecodeSetter(field, obj, value);
            }

            return obj;
        }

        public void Dispose()
        {
            _fieldCache.Clear();
        }
    }

    public static class DataObjectExtension
    {
        public static void SetDecodePropertyValue(FieldInfo fieldInfo, object obj, string rawValue)
        {
            var propertyType = fieldInfo.FieldType;

            bool parsable;
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
                if (parsable) fieldInfo.SetValue(obj, value);
            }
            else if (rawValue == "null")
            {
                parsable = true;
                fieldInfo.SetValue(obj, default);
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
}