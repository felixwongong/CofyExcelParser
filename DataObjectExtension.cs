using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
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
            _SetDecodePropertyValue(propertyInfo, ref propertyObject, kvp);
        }
        
        private static void _SetDecodePropertyValue(PropertyInfo propertyInfo, ref object propertyObject, KeyValuePair<string, object> kvp)
        {
            var fieldType = propertyInfo.PropertyType;
            var value = kvp.Value;

            bool parsable = false;
            if (value is string rawValue)
            {
                parsable = TryDecodeSingleValue(fieldType, propertyObject, rawValue);
            }
            else if (value is IList rawValues)
            {
                parsable = true;

                if (propertyInfo.GetValue(propertyObject) is not IList list)
                {
                    list = (IList)Activator.CreateInstance(fieldType);
                    propertyInfo.SetValue(propertyObject, list);
                }
                
                var separatorIndex = kvp.Key.IndexOf('.');
                if (separatorIndex == -1)
                {
                    throw new ArgumentException($"dataObject key ({kvp.Key}) is not a list element, use format: listName.fieldName");
                }
                
                var listItemType = fieldType.GetGenericArguments()[0];
                var listItemPropertyName = kvp.Key.AsSpan()[(separatorIndex + 1)..].ToString();
                var listItemProperty = listItemType.GetProperty(listItemPropertyName);
                
                if(listItemProperty == null)
                {
                    throw new KeyNotFoundException($"{listItemType} public field not found for key ({listItemPropertyName})");
                }
                
                for (var i = 0; i < rawValues.Count; i++)
                {
                    if (list.Count <= i)
                    {
                        list.Add(Activator.CreateInstance(listItemType));
                    }
                    var innerObject = list[i];
                    _SetDecodePropertyValue(listItemProperty, ref innerObject, new KeyValuePair<string, object>(listItemPropertyName, rawValues[i]));
                }
            }
            else
            {
                parsable = true;
                propertyInfo.SetValue(propertyObject, value);
            }

            if (!parsable)
            {
                throw new ArgumentException(
                    $"dataObject value ({value}) cannot parse to {propertyObject.GetType()}'s {fieldType} field {propertyInfo.Name}");
            }

            bool TryDecodeSingleValue(Type type, object obj, string raw)
            {
                if (type == typeof(bool))
                {
                    parsable = bool.TryParse(raw, out var v);
                    if (parsable) propertyInfo.SetValue(obj, v);
                }
                else if (type == typeof(int))
                {
                    parsable = int.TryParse(raw, out var v);
                    if (parsable) propertyInfo.SetValue(obj, v);
                }
                else if (type == typeof(float))
                {
                    parsable = float.TryParse(raw, out var v);
                    if (parsable) propertyInfo.SetValue(obj, v);
                }
                else if (type == typeof(double))
                {
                    parsable = double.TryParse(raw, out var v);
                    if (parsable) propertyInfo.SetValue(obj, v);
                }
                else if (type.IsEnum)
                {
                    parsable = Enum.TryParse(fieldType, raw, out var v);
                    if (parsable) propertyInfo.SetValue(obj, v);
                }
                else if (type == typeof(string))
                {
                    propertyInfo.SetValue(obj, raw);
                    parsable = true;
                }
                else if (raw == "null")
                {
                    parsable = true;
                    propertyInfo.SetValue(obj, null);
                }

                return parsable;
            }
        }
    }
}