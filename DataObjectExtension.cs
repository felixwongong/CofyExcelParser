using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CofyDev.Xml.Doc
{
    public class DataObjectEncoder: IDisposable
    {
        private Dictionary<string, FieldInfo> _fieldCache = new();

        public virtual DataObject Encode(object obj)
        {
            throw new NotImplementedException();
        }

        public virtual T DecodeAs<T>(DataObject dataObject, Action<FieldInfo, object, KeyValuePair<string, object>> propertyDecodeSetter)
        {
            var objType = typeof(T);

            var obj = Activator.CreateInstance<T>();
            if (obj == null)
            {
                throw new ArgumentException($"Cannot create instance type ({typeof(T)})");
            }

            foreach (var (key, value) in dataObject)
            {
                var fieldName = key;
                if (string.IsNullOrEmpty(fieldName))
                {
                    throw new ArgumentNullException(nameof(fieldName), "dataObject key is empty");
                }
                
                var separatorIndex = key.IndexOf('.');
                if (separatorIndex != -1)
                {
                    fieldName = key.AsSpan()[..separatorIndex].ToString();
                }

                if (!_fieldCache.TryGetValue(fieldName, out var field))
                {
                    field = objType.GetField(fieldName, BindingFlags.Instance | BindingFlags.Public);
                    _fieldCache[key] = field;
                }

                if (field == null)
                {
                    throw new KeyNotFoundException($"{objType} public field not found for key ({key})");
                }

                propertyDecodeSetter(field, obj, new KeyValuePair<string, object>(key, value));
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
        public static void SetDecodePropertyValue(FieldInfo fieldInfo, object fieldObject, KeyValuePair<string, object> kvp)
        {
            _SetDecodePropertyValue(fieldInfo, ref fieldObject, kvp);
        }
        
        private static void _SetDecodePropertyValue(FieldInfo fieldInfo, ref object fieldObject, KeyValuePair<string, object> kvp)
        {
            var fieldType = fieldInfo.FieldType;
            var key = kvp.Key;
            var value = kvp.Value;

            bool parsable = false;
            if (value is string rawValue)
            {
                parsable = TryDecodeSingleValue(fieldType, fieldObject, rawValue);
            }
            else if (value is IList rawValues)
            {
                parsable = true;

                if (fieldInfo.GetValue(fieldObject) is not IList list)
                {
                    list = (IList)Activator.CreateInstance(fieldType);
                    fieldInfo.SetValue(fieldObject, list);
                }
                
                var separatorIndex = kvp.Key.IndexOf('.');
                if (separatorIndex == -1)
                {
                    throw new ArgumentException($"dataObject key ({kvp.Key}) is not a list element, use format: listName.fieldName");
                }
                
                var listItemType = fieldType.GetGenericArguments()[0];
                var listItemFieldName = kvp.Key.AsSpan()[(separatorIndex + 1)..].ToString();
                var listItemField = listItemType.GetField(listItemFieldName);
                
                if(listItemField == null)
                {
                    throw new KeyNotFoundException($"{listItemType} public field not found for key ({listItemFieldName})");
                }
                
                for (var i = 0; i < rawValues.Count; i++)
                {
                    if (list.Count <= i)
                    {
                        list.Add(Activator.CreateInstance(listItemType));
                    }
                    var innerObject = list[i];
                    _SetDecodePropertyValue(listItemField, ref innerObject, new KeyValuePair<string, object>(listItemFieldName, rawValues[i]));
                }
            }
            else
            {
                parsable = true;
                fieldInfo.SetValue(fieldObject, value);
            }

            if (!parsable)
            {
                throw new ArgumentException(
                    $"dataObject value ({value}) cannot parse to {fieldObject.GetType()}'s {fieldType} field {fieldInfo.Name}");
            }

            bool TryDecodeSingleValue(Type type, object obj, string raw)
            {
                if (type == typeof(bool))
                {
                    parsable = bool.TryParse(raw, out var v);
                    if (parsable) fieldInfo.SetValue(obj, v);
                }
                else if (fieldType == typeof(int))
                {
                    parsable = int.TryParse(raw, out var v);
                    if (parsable) fieldInfo.SetValue(obj, v);
                }
                else if (fieldType == typeof(float))
                {
                    parsable = float.TryParse(raw, out var v);
                    if (parsable) fieldInfo.SetValue(obj, v);
                }
                else if (fieldType == typeof(double))
                {
                    parsable = double.TryParse(raw, out var v);
                    if (parsable) fieldInfo.SetValue(obj, v);
                }
                else if (fieldType.IsEnum)
                {
                    parsable = Enum.TryParse(fieldType, raw, out var v);
                    if (parsable) fieldInfo.SetValue(obj, v);
                }
                else if (fieldType == typeof(string))
                {
                    fieldInfo.SetValue(obj, raw);
                    parsable = true;
                }
                else if (raw == "null")
                {
                    parsable = true;
                    fieldInfo.SetValue(obj, null);
                }

                return parsable;
            }
        }
    }
}