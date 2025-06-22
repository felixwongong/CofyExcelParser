using System;
using System.Collections;
using System.Collections.Generic;
using cfEngine.Logging;

namespace CofyDev.Xml.Doc
{
    public partial class DataObject
    {
        public interface IValueDecoder
        {
            public Type valueType { get; }
            public bool TryDecode(object raw, Type decodedType, out object decoded);
        }

        public static class Decoder
        {
            private static Dictionary<Type, IValueDecoder> _stringDecoders = new Dictionary<Type, IValueDecoder>();
            public static IReadOnlyDictionary<Type, IValueDecoder> stringDecoders => _stringDecoders;

            static Decoder()
            {
                RegisterDecoder(new StringValueDecoder());
            }
            
            public static bool TryDecode(object raw, Type decodedType, out object decoded)
            {
                decoded = null;
                if(!TryGetDecoder(raw.GetType(), out var decoder))
                    return false;
                
                return decoder.TryDecode(raw, decodedType, out decoded);
            }
            
            public static void RegisterDecoder(IValueDecoder decoder)
            {
                _stringDecoders.Add(decoder.valueType, decoder);
            }
            
            public static bool TryGetDecoder(Type type, out IValueDecoder decoder)
            {
                return _stringDecoders.TryGetValue(type, out decoder);
            }
        }

        public class ListValueDecoder : IValueDecoder
        {
            public Type valueType => typeof(IList);
            public bool TryDecode(object raw, Type decodedType, out object decoded)
            {
                decoded = null;
                if(raw is not IList rawValues)
                    return false;
                
                if (rawValues.Count <= 0)
                    return true;
                
                var listItemType = decodedType.GetGenericArguments()[0];
                var list = (IList) Activator.CreateInstance(decodedType);
                for (var i = 0; i < rawValues.Count; i++)
                {
                    var rawValue = rawValues[i];
                    if (!Decoder.TryDecode(rawValue, decodedType, out var decodedValue) || decodedValue.GetType() != listItemType)
                    {
                        Log.LogError($"Failed to decode list item value ({rawValue}) to target type {listItemType}, index: {i}");
                        return false;
                    }
                    list.Add(decodedValue);
                }

                decoded = list;
                return true;
            }
        }
    }
}