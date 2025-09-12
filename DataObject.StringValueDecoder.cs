using System;
using System.Collections.Generic;

namespace CofyDev.Xml.Doc
{
    public partial class DataObject
    {
        public interface IStringDecoder 
        {
            public Type propertyType { get; }

            public bool TryDecode(string? raw, out object? decoded);
        }

        public class StringValueDecoder : IValueDecoder
        {
            private static Dictionary<Type, IStringDecoder> _stringDecoders = new Dictionary<Type, IStringDecoder>();
            public static IReadOnlyDictionary<Type, IStringDecoder> stringDecoders => _stringDecoders;

            public Type valueType => typeof(string);

            public StringValueDecoder()
            {
                RegisterStringDecoder(new BooleanDecoder());
                RegisterStringDecoder(new IntDecoder());
                RegisterStringDecoder(new FloatDecoder());
                RegisterStringDecoder(new DoubleDecoder());
                RegisterStringDecoder(new StringDecoder());
            }

            public bool TryDecode(object raw, Type decodedType, out object? decoded)
            {
                decoded = null;
                
                if(raw is not string rawString || !TryGetStringDecoder(decodedType, out var decoder))
                    return false;
                
                return decoder.TryDecode(rawString, out decoded);
            }
            
            public static void RegisterStringDecoder(IStringDecoder decoder)
            {
                _stringDecoders.Add(decoder.propertyType, decoder);
            }
            
            public static bool TryGetStringDecoder(Type type, out IStringDecoder decoder)
            {
                if (!_stringDecoders.TryGetValue(type, out decoder) && type.IsEnum)
                {
                    decoder = new EnumDecoder(type);
                    RegisterStringDecoder(decoder);
                }

                return decoder != null;
            }
        }
        
        public struct BooleanDecoder: IStringDecoder
        {
            public Type propertyType => typeof(bool);
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (bool.TryParse(raw, out var result))
                    return false;
                
                decoded = result;
                return true;
            }
        }
        public struct IntDecoder: IStringDecoder
        {
            public Type propertyType => typeof(int);
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (!float.TryParse(raw, out var result))
                    return false;
                
                decoded = Convert.ToInt32(result);
                return true;
            }
        }
        public struct FloatDecoder: IStringDecoder
        {
            public Type propertyType => typeof(float);
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (!float.TryParse(raw, out var result))
                    return false;
                
                decoded = result;
                return true;
            }
        }
        public struct DoubleDecoder: IStringDecoder
        {
            public Type propertyType => typeof(double);
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (!double.TryParse(raw, out var result))
                    return false;
                
                decoded = result;
                return true;
            }
        }
        public struct StringDecoder: IStringDecoder
        {
            public Type propertyType => typeof(string);
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (raw == "null")
                    return false;
                
                decoded = raw;
                return true;
            }
        }
        
        public struct EnumDecoder: IStringDecoder
        {
            private Type _enumType;
            public Type propertyType => _enumType;
            
            public EnumDecoder(Type enumType)
            {
                if (!enumType.IsEnum)
                    throw new ArgumentException($"{enumType} is not an enum type");
                
                _enumType = enumType;
            }
            
            public bool TryDecode(string? raw, out object? decoded)
            {
                decoded = null;
                if (!Enum.TryParse(_enumType, raw, out var result))
                    return false;
                
                decoded = result;
                return true;
            }
        }
    }
}