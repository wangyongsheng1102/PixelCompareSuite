using System;
using System.Globalization;
using System.IO;
using Avalonia.Data.Converters;
using Avalonia.Media.Imaging;
using Avalonia.Threading;


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
                        Bitmap? bitmap = null;
                        Dispatcher.UIThread.Invoke(() =>
                        {
                            bitmap = new Bitmap(path);
                        });
                        return bitmap;
                        // var bytes = File.ReadAllBytes(path);
                        // var ms = new MemoryStream(bytes);
                        // return new Avalonia.Media.Imaging.Bitmap(ms);
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

