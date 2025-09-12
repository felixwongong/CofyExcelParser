using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace CofyDev.Xml.Doc
{
    public class DataObjectEncoder: IDisposable
    {
        private Dictionary<string, PropertyInfo> _propertyCache = new();

        public virtual DataObject Encode(object obj)
        {
            throw new NotImplementedException();
        }

        public virtual T DecodeAs<T>(DataObject dataObject, Action<PropertyInfo, object, KeyValuePair<string, object>> propertyDecodeSetter)
        {
            var objType = typeof(T);

            var obj = Activator.CreateInstance<T>();
            if (obj == null)
            {
                throw new ArgumentException($"Cannot create instance type ({typeof(T)})");
            }

            foreach (var (key, value) in dataObject)
            {
                var propertyName = key;
                if (string.IsNullOrEmpty(propertyName))
                {
                    throw new ArgumentNullException(nameof(propertyName), "dataObject key is empty");
                }
                
                var separatorIndex = key.IndexOf('.');
                if (separatorIndex != -1)
                {
                    propertyName = key.AsSpan()[..separatorIndex].ToString();
                }

                if (!_propertyCache.TryGetValue(propertyName, out var property))
                {
                    property = objType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                    _propertyCache[key] = property;
                }

                if (property == null)
                {
                    throw new KeyNotFoundException($"{objType} public field not found for key ({key})");
                }

                propertyDecodeSetter(property, obj, new KeyValuePair<string, object>(key, value));
            }

            return obj;
        }

        public void Dispose()
        {
            _propertyCache.Clear();
        }
    }

    public static class DataObjectExtension
    {
        public static void SetDecodePropertyValue(PropertyInfo propertyInfo, object propertyObject, KeyValuePair<string, object> kvp)
        {
            _SetDecodePropertyValue(propertyInfo, propertyObject, kvp);
        }
        
        private static void _SetDecodePropertyValue(PropertyInfo propertyInfo, in object propertyObject, KeyValuePair<string, object> kvp)
        {
            var propertyType = propertyInfo.PropertyType;
            var value = kvp.Value;

            if (!DataObject.Decoder.TryDecode(value, propertyType, out var decoded))
            {
                throw new ArgumentException($"dataObject value ({value}) cannot parse to {propertyObject.GetType()}'s {propertyType} field {propertyInfo.Name}");
            }

            propertyInfo.SetValue(propertyObject, decoded);
        }
    }
}