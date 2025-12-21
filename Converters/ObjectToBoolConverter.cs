using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace PixelCompareSuite.Converters
{
    public class ObjectToBoolConverter : IValueConverter
    {
        public static readonly ObjectToBoolConverter IsNotNull = new ObjectToBoolConverter { Invert = false };
        public static readonly ObjectToBoolConverter IsNull = new ObjectToBoolConverter { Invert = true };

        public bool Invert { get; set; }

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool result = value != null;
            return Invert ? !result : result;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

