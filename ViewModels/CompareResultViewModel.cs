using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Microsoft.Office.Interop.Excel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Png;

namespace PixelCompareSuite.ViewModels
{
    public class CompareResultViewModel : INotifyPropertyChanged
    {
        private string _filePath = string.Empty;
        private string _selectedSheet = string.Empty;
        private string _column1 = "B";
        private string _column2 = "P";
        private int _currentPage = 1;
        private int _pageSize = 10;
        private int _totalItems = 0;
        private CompareItemViewModel? _selectedItem;
        private double _progress = 0;
        private string _statusMessage = "就绪";
        private bool _isProcessing = false;
        private TopLevel? _topLevel;

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                if (_selectedSheet != value)
                {
                    _selectedSheet = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Column1
        {
            get => _column1;
            set
            {
                if (_column1 != value)
                {
                    _column1 = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Column2
        {
            get => _column2;
            set
            {
                if (_column2 != value)
                {
                    _column2 = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<string> AvailableSheets { get; } = new ObservableCollection<string>();
        public ObservableCollection<CompareItemViewModel> CompareItems { get; } = new ObservableCollection<CompareItemViewModel>();
        public ObservableCollection<CompareItemViewModel> CurrentPageItems { get; } = new ObservableCollection<CompareItemViewModel>();

        public int CurrentPage
        {
            get => _currentPage;
            set
            {
                if (_currentPage != value && value > 0)
                {
                    _currentPage = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TotalPages));
                    OnPropertyChanged(nameof(CanGoToPreviousPage));
                    OnPropertyChanged(nameof(CanGoToNextPage));
                    UpdateCurrentPageItems();
                }
            }
        }

        public int PageSize
        {
            get => _pageSize;
            set
            {
                if (_pageSize != value && value > 0)
                {
                    _pageSize = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TotalPages));
                    UpdateCurrentPageItems();
                }
            }
        }

        public int TotalItems
        {
            get => _totalItems;
            set
            {
                if (_totalItems != value)
                {
                    _totalItems = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TotalPages));
                }
            }
        }

        public int TotalPages => (int)Math.Ceiling((double)TotalItems / PageSize);

        public bool CanGoToPreviousPage => CurrentPage > 1;
        public bool CanGoToNextPage => CurrentPage < TotalPages;

        public CompareItemViewModel? SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem != value)
                {
                    _selectedItem = value;
                    OnPropertyChanged();
                }
            }
        }

        public double Progress
        {
            get => _progress;
            set
            {
                if (Math.Abs(_progress - value) > 0.01)
                {
                    _progress = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set
            {
                if (_isProcessing != value)
                {
                    _isProcessing = value;
                    OnPropertyChanged();
                }
            }
        }

        public ICommand SelectFileCommand { get; }
        public ICommand LoadDataCommand { get; }
        public ICommand PreviousPageCommand { get; }
        public ICommand NextPageCommand { get; }
        public ICommand SelectItemCommand { get; }

        public CompareResultViewModel()
        {
            SelectFileCommand = new RelayCommand(async () => await SelectFileAsync());
            var loadCommand = new RelayCommand(async () => await LoadDataAsync(), () => !string.IsNullOrEmpty(FilePath) && !string.IsNullOrEmpty(SelectedSheet));
            LoadDataCommand = loadCommand;
            
            // 当 FilePath 或 SelectedSheet 改变时，更新命令状态
            PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(FilePath) || e.PropertyName == nameof(SelectedSheet))
                {
                    loadCommand.RaiseCanExecuteChanged();
                }
            };
            
            PreviousPageCommand = new RelayCommand(() => CurrentPage--, () => CanGoToPreviousPage);
            NextPageCommand = new RelayCommand(() => CurrentPage++, () => CanGoToNextPage);
            SelectItemCommand = new RelayCommand<CompareItemViewModel>(async item => 
            {
                SelectedItem = item;
                if (item != null && !item.IsComparisonLoaded)
                {
                    await LoadComparisonForItem(item);
                }
            });
        }

        public void SetTopLevel(TopLevel topLevel)
        {
            _topLevel = topLevel;
        }

        private async Task SelectFileAsync()
        {
            if (_topLevel == null) return;

            var files = await _topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "选择 Excel 文件",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    FilePickerFileTypes.All,
                    new FilePickerFileType("Excel 文件")
                    {
                        Patterns = new[] { "*.xlsx", "*.xls" },
                        MimeTypes = new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel" }
                    }
                }
            });

            if (files.Count > 0 && files[0].Path.IsFile)
            {
                FilePath = files[0].Path.LocalPath;
                await LoadSheetNamesAsync();
                ((RelayCommand)LoadDataCommand).RaiseCanExecuteChanged();
            }
        }

        private async Task LoadSheetNamesAsync()
        {
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
                return;

            try
            {
                await Task.Run(() =>
                {
                    Application? excelApp = null;
                    Workbook? workbook = null;
                    try
                    {
                        excelApp = new Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;
                        
                        workbook = excelApp.Workbooks.Open(FilePath, ReadOnly: true);
                        var sheetNames = new List<string>();
                        
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            sheetNames.Add(sheet.Name);
                        }
                        
                        Dispatcher.UIThread.Post(() =>
                        {
                            AvailableSheets.Clear();
                            foreach (var name in sheetNames)
                            {
                                AvailableSheets.Add(name);
                            }
                            if (AvailableSheets.Count > 0 && string.IsNullOrEmpty(SelectedSheet))
                            {
                                SelectedSheet = AvailableSheets[0];
                            }
                        });
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"读取 Sheet 列表失败: {ex.Message}";
            }
        }

        private async Task LoadDataAsync()
        {
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
            {
                StatusMessage = "文件不存在";
                return;
            }

            if (string.IsNullOrEmpty(SelectedSheet))
            {
                StatusMessage = "请选择 Sheet 页";
                return;
            }

            IsProcessing = true;
            StatusMessage = "正在加载数据...";
            Progress = 0;

            try
            {
                CompareItems.Clear();
                SelectedItem = null;

                await Task.Run(() =>
                {
                    Application? excelApp = null;
                    Workbook? workbook = null;
                    Worksheet? worksheet = null;
                    try
                    {
                        excelApp = new Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;
                        
                        workbook = excelApp.Workbooks.Open(FilePath, ReadOnly: true);
                        
                        // 查找指定的工作表
                        worksheet = null;
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            if (sheet.Name == SelectedSheet)
                            {
                                worksheet = sheet;
                                break;
                            }
                        }
                        
                        if (worksheet == null)
                        {
                            Dispatcher.UIThread.Post(() =>
                            {
                                StatusMessage = $"Sheet '{SelectedSheet}' 不存在";
                            });
                            return;
                        }

                        var column1Index = GetColumnIndex(Column1);
                        var column2Index = GetColumnIndex(Column2);
                        
                        // 获取使用的行数
                        Range? usedRange = worksheet.UsedRange;
                        int rowCount = usedRange != null ? usedRange.Rows.Count : 0;
                        var items = new List<CompareItemViewModel>();

                        for (int row = 2; row <= rowCount; row++) // 从第2行开始，假设第1行是标题
                        {
                            Range? cell1 = worksheet.Cells[row, column1Index];
                            Range? cell2 = worksheet.Cells[row, column2Index];
                            
                            var image1Path = cell1?.Value2?.ToString() ?? string.Empty;
                            var image2Path = cell2?.Value2?.ToString() ?? string.Empty;

                            // 释放 COM 对象
                            if (cell1 != null) Marshal.ReleaseComObject(cell1);
                            if (cell2 != null) Marshal.ReleaseComObject(cell2);

                            // 如果两列都有值，则创建对比项
                            if (!string.IsNullOrWhiteSpace(image1Path) && !string.IsNullOrWhiteSpace(image2Path))
                            {
                                items.Add(new CompareItemViewModel
                                {
                                    RowIndex = row,
                                    Image1Path = image1Path.Trim(),
                                    Image2Path = image2Path.Trim()
                                });
                            }

                            // 更新进度
                            var progress = (double)(row - 1) / rowCount * 50; // 前50%用于读取数据
                            Dispatcher.UIThread.Post(() =>
                            {
                                Progress = progress;
                                StatusMessage = $"正在读取第 {row} 行...";
                            });
                        }

                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                        if (worksheet != null) Marshal.ReleaseComObject(worksheet);

                        Dispatcher.UIThread.Post(() =>
                        {
                            foreach (var item in items)
                            {
                                CompareItems.Add(item);
                            }
                            TotalItems = CompareItems.Count;
                            CurrentPage = 1;
                            UpdateCurrentPageItems();
                            StatusMessage = $"已加载 {TotalItems} 个对比项";
                            Progress = 50;
                        });
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                });

                StatusMessage = $"已加载 {TotalItems} 个对比项，点击列表项查看对比结果";
                Progress = 100;
            }
            catch (Exception ex)
            {
                StatusMessage = $"加载失败: {ex.Message}";
                Progress = 0;
            }
            finally
            {
                IsProcessing = false;
            }
        }

        private int GetColumnIndex(string columnName)
        {
            int index = 0;
            foreach (char c in columnName.ToUpper())
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index;
        }

        private async Task LoadComparisonForItem(CompareItemViewModel item)
        {
            if (item == null || item.IsComparisonLoaded) return;

            try
            {
                StatusMessage = $"正在对比行 {item.RowIndex} 的图片...";
                item.IsLoading = true;

                await Task.Run(() =>
                {
                    var image1Path = item.Image1Path;
                    var image2Path = item.Image2Path;

                    // 检查文件是否存在
                    if (!File.Exists(image1Path) || !File.Exists(image2Path))
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            item.DifferencePercentage = -1;
                            item.IsComparisonLoaded = true;
                            item.IsLoading = false;
                            StatusMessage = $"行 {item.RowIndex}: 图片文件不存在";
                        });
                        return;
                    }

                    try
                    {
                        // 使用 ImageSharp 进行像素对比
                        using var img1 = Image.Load<Rgba32>(image1Path);
                        using var img2 = Image.Load<Rgba32>(image2Path);

                        // 检查图片尺寸是否一致
                        bool isSizeMatch = img1.Width == img2.Width && img1.Height == img2.Height;
                        string sizeInfo = $"图1: {img1.Width}x{img1.Height}, 图2: {img2.Width}x{img2.Height}";

                        // 如果尺寸不一致，不进行对比，只标识出来
                        if (!isSizeMatch)
                        {
                            Dispatcher.UIThread.Post(() =>
                            {
                                item.DifferencePercentage = -2; // 使用 -2 表示尺寸不一致
                                item.IsSizeMismatch = true;
                                item.SizeInfo = sizeInfo;
                                item.Image1BitmapPath = image1Path;
                                item.Image2BitmapPath = image2Path;
                                item.IsComparisonLoaded = true;
                                item.IsLoading = false;
                                StatusMessage = $"行 {item.RowIndex}: 图片尺寸不一致 - {sizeInfo}";
                            });
                            return;
                        }

                        // 尺寸一致，进行对比
                        // 克隆图片用于处理
                        using var img1Clone = img1.Clone();
                        using var img2Clone = img2.Clone();

                        // 转换为灰度图进行对比
                        img1Clone.Mutate(x => x.Grayscale());
                        img2Clone.Mutate(x => x.Grayscale());

                        // 计算差异
                        var width = img1.Width;
                        var height = img1.Height;
                        var totalPixels = width * height;
                        var threshold = 30; // 差异阈值
                        var differentPixels = 0;
                        var diffMap = new bool[width, height]; // 记录差异位置

                        // 遍历像素计算差异
                        for (int y = 0; y < height; y++)
                        {
                            for (int x = 0; x < width; x++)
                            {
                                var pixel1 = img1Clone[x, y];
                                var pixel2 = img2Clone[x, y];
                                
                                // 计算灰度值差异
                                var diff = Math.Abs(pixel1.R - pixel2.R);
                                
                                if (diff > threshold)
                                {
                                    differentPixels++;
                                    diffMap[x, y] = true;
                                }
                                else
                                {
                                    diffMap[x, y] = false;
                                }
                            }
                        }

                        var differencePercentage = (double)differentPixels / totalPixels * 100;

                        // 生成差异可视化图像（彩色）
                        using var diffImage = img1.Clone();
                        diffImage.Mutate(x =>
                        {
                            for (int y = 0; y < height; y++)
                            {
                                for (int px = 0; px < width; px++)
                                {
                                    if (diffMap[px, y])
                                    {
                                        var pixel1 = img1[px, y];
                                        var pixel2 = img2[px, y];
                                        
                                        // 计算差异并增强显示
                                        var r = Math.Min(255, Math.Abs(pixel1.R - pixel2.R) * 3);
                                        var g = Math.Min(255, Math.Abs(pixel1.G - pixel2.G) * 3);
                                        var b = Math.Min(255, Math.Abs(pixel1.B - pixel2.B) * 3);
                                        
                                        diffImage[px, y] = new Rgba32((byte)r, (byte)g, (byte)b, 255);
                                    }
                                }
                            }
                        });

                        // 找到差异区域并绘制红框
                        var minArea = 50; // 最小区域面积
                        var differenceRegions = FindDifferenceRegions(diffMap, width, height, minArea);

                        // 在原图1上标记差异区域
                        using var markedImage1 = img1.Clone();
                        markedImage1.Mutate(x =>
                        {
                            var redPen = Pens.Solid(Color.Red, 2f);
                            foreach (var rect in differenceRegions)
                            {
                                // 绘制矩形边框
                                var topLeft = new SixLabors.ImageSharp.PointF(rect.X, rect.Y);
                                var topRight = new SixLabors.ImageSharp.PointF(rect.X + rect.Width, rect.Y);
                                var bottomRight = new SixLabors.ImageSharp.PointF(rect.X + rect.Width, rect.Y + rect.Height);
                                var bottomLeft = new SixLabors.ImageSharp.PointF(rect.X, rect.Y + rect.Height);
                                
                                x.DrawLines(redPen, topLeft, topRight, bottomRight, bottomLeft, topLeft);
                            }
                        });

                        // 在原图2上标记差异区域
                        using var markedImage2 = img2.Clone();
                        markedImage2.Mutate(x =>
                        {
                            var redPen = Pens.Solid(Color.Red, 2f);
                            foreach (var rect in differenceRegions)
                            {
                                // 绘制矩形边框
                                var topLeft = new SixLabors.ImageSharp.PointF(rect.X, rect.Y);
                                var topRight = new SixLabors.ImageSharp.PointF(rect.X + rect.Width, rect.Y);
                                var bottomRight = new SixLabors.ImageSharp.PointF(rect.X + rect.Width, rect.Y + rect.Height);
                                var bottomLeft = new SixLabors.ImageSharp.PointF(rect.X, rect.Y + rect.Height);
                                
                                x.DrawLines(redPen, topLeft, topRight, bottomRight, bottomLeft, topLeft);
                            }
                        });

                        // 保存标记后的图片和差异图到临时文件
                        var tempDir = Path.Combine(Path.GetTempPath(), "PixelCompareSuite");
                        Directory.CreateDirectory(tempDir);
                        var guid = Guid.NewGuid().ToString("N");
                        var diffImagePath = Path.Combine(tempDir, $"diff_{item.RowIndex}_{guid}.png");
                        var markedImage1Path = Path.Combine(tempDir, $"marked1_{item.RowIndex}_{guid}.png");
                        var markedImage2Path = Path.Combine(tempDir, $"marked2_{item.RowIndex}_{guid}.png");
                        
                        await diffImage.SaveAsync(diffImagePath, new PngEncoder());
                        await markedImage1.SaveAsync(markedImage1Path, new PngEncoder());
                        await markedImage2.SaveAsync(markedImage2Path, new PngEncoder());

                        Dispatcher.UIThread.Post(() =>
                        {
                            item.DifferencePercentage = differencePercentage;
                            item.DifferenceImagePath = diffImagePath;
                            // 使用标记后的图片
                            item.Image1BitmapPath = markedImage1Path;
                            item.Image2BitmapPath = markedImage2Path;
                            item.IsSizeMismatch = false;
                            item.SizeInfo = sizeInfo;
                            item.IsComparisonLoaded = true;
                            item.IsLoading = false;
                            StatusMessage = $"行 {item.RowIndex} 对比完成，差异度: {differencePercentage:F2}%";
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            item.DifferencePercentage = -1;
                            item.IsComparisonLoaded = true;
                            item.IsLoading = false;
                            StatusMessage = $"行 {item.RowIndex} 对比失败: {ex.Message}";
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                item.IsLoading = false;
                StatusMessage = $"处理失败: {ex.Message}";
            }
        }

        private void UpdateCurrentPageItems()
        {
            CurrentPageItems.Clear();
            var startIndex = (CurrentPage - 1) * PageSize;
            var items = CompareItems.Skip(startIndex).Take(PageSize);
            foreach (var item in items)
            {
                CurrentPageItems.Add(item);
            }
        }

        private List<SixLabors.ImageSharp.Rectangle> FindDifferenceRegions(bool[,] diffMap, int width, int height, int minArea)
        {
            var regions = new List<SixLabors.ImageSharp.Rectangle>();
            var visited = new bool[width, height];

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    if (diffMap[x, y] && !visited[x, y])
                    {
                        // 使用 BFS 找到连通区域
                        var region = FindConnectedRegion(diffMap, visited, width, height, x, y);
                        
                        if (region.Width * region.Height >= minArea)
                        {
                            regions.Add(region);
                        }
                    }
                }
            }

            return regions;
        }

        private SixLabors.ImageSharp.Rectangle FindConnectedRegion(bool[,] diffMap, bool[,] visited, int width, int height, int startX, int startY)
        {
            var minX = startX;
            var maxX = startX;
            var minY = startY;
            var maxY = startY;

            var queue = new Queue<(int x, int y)>();
            queue.Enqueue((startX, startY));
            visited[startX, startY] = true;

            while (queue.Count > 0)
            {
                var (x, y) = queue.Dequeue();

                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);

                // 检查四个方向的邻居
                var directions = new[] { (0, 1), (0, -1), (1, 0), (-1, 0) };
                foreach (var (dx, dy) in directions)
                {
                    var nx = x + dx;
                    var ny = y + dy;

                    if (nx >= 0 && nx < width && ny >= 0 && ny < height &&
                        diffMap[nx, ny] && !visited[nx, ny])
                    {
                        visited[nx, ny] = true;
                        queue.Enqueue((nx, ny));
                    }
                }
            }

            return new SixLabors.ImageSharp.Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class CompareItemViewModel : INotifyPropertyChanged
    {
        private int _rowIndex;
        private string _image1Path = string.Empty;
        private string _image2Path = string.Empty;
        private double _differencePercentage;
        private string? _differenceImagePath;
        private string? _image1BitmapPath;
        private string? _image2BitmapPath;
        private bool _isComparisonLoaded = false;
        private bool _isLoading = false;
        private bool _isSizeMismatch = false;
        private string _sizeInfo = string.Empty;

        public int RowIndex
        {
            get => _rowIndex;
            set
            {
                if (_rowIndex != value)
                {
                    _rowIndex = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Image1Path
        {
            get => _image1Path;
            set
            {
                if (_image1Path != value)
                {
                    _image1Path = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Image2Path
        {
            get => _image2Path;
            set
            {
                if (_image2Path != value)
                {
                    _image2Path = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? Image1BitmapPath
        {
            get => _image1BitmapPath;
            set
            {
                if (_image1BitmapPath != value)
                {
                    _image1BitmapPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? Image2BitmapPath
        {
            get => _image2BitmapPath;
            set
            {
                if (_image2BitmapPath != value)
                {
                    _image2BitmapPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public double DifferencePercentage
        {
            get => _differencePercentage;
            set
            {
                if (Math.Abs(_differencePercentage - value) > 0.01)
                {
                    _differencePercentage = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DifferencePercentageText));
                }
            }
        }

        public string DifferencePercentageText
        {
            get
            {
                if (IsSizeMismatch)
                    return "尺寸不一致";
                if (DifferencePercentage < 0)
                    return "对比失败";
                return $"{DifferencePercentage:F2}%";
            }
        }

        public string? DifferenceImagePath
        {
            get => _differenceImagePath;
            set
            {
                if (_differenceImagePath != value)
                {
                    _differenceImagePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsComparisonLoaded
        {
            get => _isComparisonLoaded;
            set
            {
                if (_isComparisonLoaded != value)
                {
                    _isComparisonLoaded = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsSizeMismatch
        {
            get => _isSizeMismatch;
            set
            {
                if (_isSizeMismatch != value)
                {
                    _isSizeMismatch = value;
                    OnPropertyChanged();
                }
            }
        }

        public string SizeInfo
        {
            get => _sizeInfo;
            set
            {
                if (_sizeInfo != value)
                {
                    _sizeInfo = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool>? _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke((T)parameter!) ?? true;

        public void Execute(object? parameter) => _execute((T)parameter!);

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}

