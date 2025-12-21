using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace PixelCompareSuite.Converters
{
    public class PathToBitmapConverter : IValueConverter
    {
        public static readonly PathToBitmapConverter Instance = new PathToBitmapConverter();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is string path && !string.IsNullOrWhiteSpace(path))
            {
                try
                {
                    if (System.IO.File.Exists(path))
                    {
                        return new Bitmap(path);
                    }
                }
                catch
                {
                    // 如果加载失败，返回 null
                }
            }
            return null;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

