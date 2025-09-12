using System;
using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace CofyDev.Xml.Doc
{
    public partial class DataObject
    {
        public interface IValueDecoder
        {
            public Type valueType { get; }
            public bool TryDecode(object raw, Type decodedType, out object? decoded);
        }

        public static class Decoder
        {
            private static List<IValueDecoder> _decoders = new();

            static Decoder()
            {
                RegisterDecoder(new StringValueDecoder());
                RegisterDecoder(new ArrayDecoder());
            }
            
            public static bool TryDecode(object raw, Type decodedType, out object? decoded)
            {
                decoded = null;
                if(!TryGetDecoder(raw.GetType(), out var decoder))
                    return false;
                
                return decoder.TryDecode(raw, decodedType, out decoded);
            }
            
            public static void RegisterDecoder(IValueDecoder decoder)
            {
                _decoders.Add(decoder);
            }
            
            public static bool TryGetDecoder(Type type, out IValueDecoder decoder)
            {
                decoder = null;
                try
                {
                    decoder = _decoders.First(x => x.valueType.IsAssignableFrom(type));
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine(ex);
                    return false;
                }

                return true;
            }
        }

        public class ArrayDecoder : IValueDecoder
        {
            public Type valueType => typeof(Array);
            public bool TryDecode(object raw, Type decodedType, out object? decoded)
            {
                if (raw is not Array rawValues || rawValues.Length <= 0)
                {
                    decoded = Array.Empty<object>();
                    return false;
                }

                var elementType = decodedType.IsArray ? decodedType.GetElementType() : typeof(object);
                if (elementType == null)
                {
                    decoded = Array.Empty<object>();
                    return false;
                }

                if(Activator.CreateInstance(decodedType, rawValues.Length) is not Array array)
                {
                    decoded = Array.Empty<object>();
                    return false;
                }
                
                for (int i = 0; i < rawValues.Length; i++)
                {
                    var rawValue = rawValues.GetValue(i);
                    if (rawValue != null && Decoder.TryDecode(rawValue, elementType, out var elementDecoded))
                    {
                        array.SetValue(elementDecoded, i);
                    }
                }

                decoded = array;
                return true;
            }
        }
    }
}