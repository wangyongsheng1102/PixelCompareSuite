using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using OfficeOpenXml;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Drawing;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using ImageSharp = SixLabors.ImageSharp; // 使用别名避免与 Avalonia.Controls.Image 冲突
// using DrawingPath = System.IO.Path;
using IOPath = System.IO.Path;
using DrawingPath = SixLabors.ImageSharp.Drawing.Path;

using OfficeOpenXml.Drawing;

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
        private string _statusMessage = "準備できました";
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
        public ICommand ExportHtmlReportCommand { get; }

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
            
            var exportCommand = new RelayCommand(async () => await ExportHtmlReportAsync(), 
                () => !string.IsNullOrEmpty(FilePath) && File.Exists(FilePath));

            ExportHtmlReportCommand = exportCommand;
            
            PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(FilePath))
                {
                    exportCommand.RaiseCanExecuteChanged();
                }
            };
            
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
                Title = "Excelファイルを選択",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    FilePickerFileTypes.All,
                    new FilePickerFileType("Excelファイル")
                    {
                        Patterns = new[] { "*.xlsx", "*.xls", "*.xlsm" },
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
                    // ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage.License.SetNonCommercialPersonal("JiMmY");
                    
                    using (var package = new ExcelPackage(new FileInfo(FilePath)))
                    {
                        var sheetNames = new List<string>();
                        
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            var column1Index = GetColumnIndex(Column1);
                            var column2Index = GetColumnIndex(Column2);
                            
                            // 获取使用的行数
                            int rowCount = GetActualLastRow(worksheet);
                            bool sheetFlag = false;
                            
                            for (int row = 2; row <= rowCount; row++) // 从第2行开始，假设第1行是标题
                            {
                                
                                var pic1 = GetPictureAtCell(worksheet, row, column1Index);
                                var pic2 = GetPictureAtCell(worksheet, row, column2Index);

                                if (pic1 != null && pic2 != null)
                                {
                                    sheetFlag = true;
                                    break;
                                }

                            }
                            
                            if (sheetFlag) sheetNames.Add(worksheet.Name);
                            
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
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"Excelシートリストの取得に失敗しました: {ex.Message}";
            }
        }

        private async Task LoadDataAsync()
        {
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
            {
                StatusMessage = "ファイルは存在していません。";
                return;
            }

            if (string.IsNullOrEmpty(SelectedSheet))
            {
                StatusMessage = "シートを選択してください。";
                return;
            }

            IsProcessing = true;
            StatusMessage = "データを読み込み中です...";
            Progress = 0;

            try
            {
                CompareItems.Clear();
                SelectedItem = null;
                
                var items = new List<CompareItemViewModel>();
                
                await Task.Run(() =>
                {
                    // ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage.License.SetNonCommercialPersonal("JiMmY");
                    
                    using (var package = new ExcelPackage(new FileInfo(FilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[SelectedSheet];
                        
                        if (worksheet == null)
                        {
                            Dispatcher.UIThread.Post(() =>
                            {
                                StatusMessage = $"シート '{SelectedSheet}' は存在していません。";
                            });
                            return;
                        }

                        var column1Index = GetColumnIndex(Column1);
                        var column2Index = GetColumnIndex(Column2);
                        
                        // 获取使用的行数
                        int rowCount = GetActualLastRow(worksheet);
                        // var items = new List<CompareItemViewModel>();
                        
                        for (int row = 2; row <= rowCount; row++) // 从第2行开始，假设第1行是标题
                        {
                            // var cell1 = worksheet.Cells[row, column1Index];
                            // var cell2 = worksheet.Cells[row, column2Index];
                            //
                            // var image1Path = cell1?.Text ?? string.Empty;
                            // var image2Path = cell2?.Text ?? string.Empty;
                            
                            var pic1 = GetPictureAtCell(worksheet, row, column1Index);
                            var pic2 = GetPictureAtCell(worksheet, row, column2Index);

                            if (pic1 != null && pic2 != null)
                            {
                                var image1Path = SavePictureToTempFile(pic1, row, "c1");
                                var image2Path = SavePictureToTempFile(pic2, row, "c2");
                                items.Add(new CompareItemViewModel
                                {
                                    RowIndex = row,
                                    Image1Path = image1Path.Trim(),
                                    Image2Path = image2Path.Trim()
                                });
                            }
                            
                            // 更新进度
                            var progress = (double)(row - 1) / rowCount * 100; // 前50%用于读取数据
                            Dispatcher.UIThread.Post(() =>
                            {
                                Progress = progress;
                                StatusMessage = $"第 {row} 行を読み込み中です...";
                            });
                        }

                        Dispatcher.UIThread.Post(() =>
                        {
                            foreach (var item in items)
                            {
                                CompareItems.Add(item);
                            }
                            TotalItems = CompareItems.Count;
                            CurrentPage = 1;
                            UpdateCurrentPageItems();
                            StatusMessage = $" {TotalItems} 件読み込みました。";
                            Progress = 100;
                        });
                        
                        
                    }
                });

                StatusMessage = $" {TotalItems} 件を読み込みました、一覧の項目をクリックして比較結果を表示。";
                Progress = 100;
                
                await LoadComparisonForAllItems(items);
                
            }
            catch (Exception ex)
            {
                StatusMessage = $"読み込みことは失敗しました: {ex.Message}";
                Progress = 0;
            }
            finally
            {
                IsProcessing = false;
            }
        }
        
        private async Task LoadComparisonForAllItems(IEnumerable<CompareItemViewModel> items)
        {
            if (items == null || !items.Any())
                return;
    
            // 获取总数量用于进度报告
            var itemList = items.ToList();
            int total = itemList.Count;
            int completed = 0;
    
            // 控制并发数
            var semaphore = new SemaphoreSlim(5); // 最多同时处理5个
            var tasks = new List<Task>();
    
            foreach (var item in itemList)
            {
                await semaphore.WaitAsync();
        
                var task = Task.Run(async () =>
                {
                    try
                    {
                        await LoadComparisonForItem(item);
                    }
                    finally
                    {
                        semaphore.Release();
                
                        // 更新进度
                        completed++;
                        await Dispatcher.UIThread.InvokeAsync(() =>
                        {
                            Progress = 100 * completed / total;
                            StatusMessage = $"比較処理中... {completed}/{total}";
                        });
                    }
                });
        
                tasks.Add(task);
            }
    
            await Task.WhenAll(tasks);
            StatusMessage = "すべての比較データを読み込み完了しました";
        }

        private ExcelPicture? GetPictureAtCell(
            ExcelWorksheet worksheet,
            int rowIndex,
            int columnIndex)
        {
            foreach (var drawing in worksheet.Drawings)
            {
                if (drawing is ExcelPicture pic)
                {
                    if (pic.From.Row == rowIndex -1 &&
                        pic.From.Column == columnIndex -1)
                    {
                        return pic;
                    }
                }
            }

            return null;
        }

        private string SavePictureToTempFile(
            ExcelPicture picture,
            int rowIndex,
            string suffix)
        {
            var tempDir = IOPath.Combine(
                IOPath.GetTempPath(),
                "PixelCompareSuite",
                "excel_images"
            );
            
            Directory.CreateDirectory(tempDir);
            
            var filePath = IOPath.Combine(tempDir, $"row{rowIndex}_{suffix}_{Guid.NewGuid():N}.png");
            
            File.WriteAllBytes(filePath,picture.Image.ImageBytes);

            return filePath;
        }

        private int GetActualLastRow(
            ExcelWorksheet worksheet)
        {
            int lastRowFromCells = worksheet.Dimension?.End.Row ?? 0;

            int lastRowFromPictures = 0;

            foreach (var d in worksheet.Drawings)
            {
                if (d is ExcelPicture pic)
                {
                    lastRowFromPictures = Math.Max(
                        lastRowFromPictures,
                        Math.Max(pic.From.Row, pic.To.Row)+1);
                }
            }
            return Math.Max(lastRowFromCells, lastRowFromPictures);
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
                StatusMessage = $" {item.RowIndex} 行目の画像を比較中です...";
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
                            StatusMessage = $"行目 {item.RowIndex}: 画像ファイルは存在していません。";
                        });
                        return;
                    }

                    try
                    {
                        // 使用 ImageSharp 进行像素对比
                        using var img1 = ImageSharp.Image.Load<Rgba32>(image1Path);
                        using var img2 = ImageSharp.Image.Load<Rgba32>(image2Path);

                        // 检查图片尺寸是否一致
                        bool isSizeMatch = img1.Width == img2.Width && img1.Height == img2.Height;
                        string sizeInfo = $"画像①: {img1.Width}x{img1.Height}, 画像②: {img2.Width}x{img2.Height}";

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
                                StatusMessage = $"行目 {item.RowIndex}: 画像のピクセルは不一致です - {sizeInfo}";
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
                        using var diffImage = img2.Clone();
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
                        var minArea = 50; // 最小区域面积（像素数）
                        var mergeDistance = 10; // 合并距离阈值（像素）
                        var expandPixels = 3; // 膨胀像素数
                        
                        // 先进行形态学膨胀操作，连接相近的差异点
                        var expandedDiffMap = ExpandDiffMap(diffMap, width, height, expandPixels);
                        
                        // 找到差异区域
                        var differenceRegions = FindDifferenceRegions(expandedDiffMap, width, height, minArea);
                        
                        // 合并相近的区域
                        differenceRegions = MergeNearbyRegions(differenceRegions, mergeDistance);

                        // 在原图1上标记差异区域（使用像素操作绘制红色矩形边框）
                        using var markedImage1 = img1.Clone();
                        // DrawRectangles(markedImage1, differenceRegions);
                        DrawRectanglesWithIndexOutside_NoMeasure(markedImage1, differenceRegions);

                        // 在原图2上标记差异区域
                        using var markedImage2 = img2.Clone();
                        // DrawRectangles(markedImage2, differenceRegions);
                        DrawRectanglesWithIndexOutside_NoMeasure(markedImage2, differenceRegions);

                        // 保存标记后的图片和差异图到临时文件
                        var tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PixelCompareSuite");
                        Directory.CreateDirectory(tempDir);
                        var guid = Guid.NewGuid().ToString("N");
                        var diffImagePath = System.IO.Path.Combine(tempDir, $"diff_{item.RowIndex}_{guid}.png");
                        var markedImage1Path = System.IO.Path.Combine(tempDir, $"marked1_{item.RowIndex}_{guid}.png");
                        var markedImage2Path = System.IO.Path.Combine(tempDir, $"marked2_{item.RowIndex}_{guid}.png");
                        
                        // 使用同步方法保存图片（在 Task.Run 中）
                        diffImage.Save(diffImagePath, new PngEncoder());
                        markedImage1.Save(markedImage1Path, new PngEncoder());
                        markedImage2.Save(markedImage2Path, new PngEncoder());

                        Dispatcher.UIThread.Post(() =>
                        {
                            item.DiffCount = differenceRegions.Count;
                            item.DifferencePercentage = differencePercentage;
                            item.DifferenceImagePath = diffImagePath;
                            // 使用标记后的图片
                            item.Image1BitmapPath = markedImage1Path;
                            item.Image2BitmapPath = markedImage2Path;
                            item.IsSizeMismatch = false;
                            item.SizeInfo = sizeInfo;
                            item.IsComparisonLoaded = true;
                            item.IsLoading = false;
                            StatusMessage = $"行目 {item.RowIndex} 比較完了です、差異度: {differencePercentage:F2}%、差異数：{differenceRegions.Count}";
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            item.DifferencePercentage = -1;
                            item.IsComparisonLoaded = true;
                            item.IsLoading = false;
                            StatusMessage = $"行目 {item.RowIndex} 比較失敗しました: {ex.Message}";
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                item.IsLoading = false;
                StatusMessage = $"処理失敗: {ex.Message}";
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

        // 膨胀差异图，连接相近的差异点
        private bool[,] ExpandDiffMap(bool[,] diffMap, int width, int height, int expandPixels)
        {
            var expanded = new bool[width, height];
            
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    if (diffMap[x, y])
                    {
                        // 在周围 expandPixels 范围内标记为差异
                        for (int dy = -expandPixels; dy <= expandPixels; dy++)
                        {
                            for (int dx = -expandPixels; dx <= expandPixels; dx++)
                            {
                                int nx = x + dx;
                                int ny = y + dy;
                                if (nx >= 0 && nx < width && ny >= 0 && ny < height)
                                {
                                    expanded[nx, ny] = true;
                                }
                            }
                        }
                    }
                }
            }
            
            return expanded;
        }

        // 合并相近的矩形区域
        private List<SixLabors.ImageSharp.Rectangle> MergeNearbyRegions(List<SixLabors.ImageSharp.Rectangle> regions, int mergeDistance)
        {
            if (regions.Count <= 1) return regions;
            
            var merged = new List<SixLabors.ImageSharp.Rectangle>();
            var used = new bool[regions.Count];
            
            for (int i = 0; i < regions.Count; i++)
            {
                if (used[i]) continue;
                
                var current = regions[i];
                used[i] = true;
                
                // 查找所有与当前区域相近的区域并合并
                bool foundNearby = true;
                while (foundNearby)
                {
                    foundNearby = false;
                    for (int j = i + 1; j < regions.Count; j++)
                    {
                        if (used[j]) continue;
                        
                        var other = regions[j];
                        
                        // 计算两个矩形的最小距离
                        int distance = CalculateRectDistance(current, other);
                        
                        if (distance <= mergeDistance)
                        {
                            // 合并两个矩形
                            int minX = Math.Min(current.X, other.X);
                            int minY = Math.Min(current.Y, other.Y);
                            int maxX = Math.Max(current.X + current.Width, other.X + other.Width);
                            int maxY = Math.Max(current.Y + current.Height, other.Y + other.Height);
                            
                            current = new SixLabors.ImageSharp.Rectangle(minX, minY, maxX - minX, maxY - minY);
                            used[j] = true;
                            foundNearby = true;
                        }
                    }
                }
                
                merged.Add(current);
            }
            
            return merged;
        }

        // 计算两个矩形之间的最小距离
        private int CalculateRectDistance(SixLabors.ImageSharp.Rectangle r1, SixLabors.ImageSharp.Rectangle r2)
        {
            int r1Right = r1.X + r1.Width;
            int r1Bottom = r1.Y + r1.Height;
            int r2Right = r2.X + r2.Width;
            int r2Bottom = r2.Y + r2.Height;
            
            // 如果两个矩形重叠或相邻，返回0
            if (r1.X <= r2Right && r2.X <= r1Right && r1.Y <= r2Bottom && r2.Y <= r1Bottom)
                return 0;
            
            // 计算最小距离
            int dx = Math.Max(0, Math.Max(r1.X - r2Right, r2.X - r1Right));
            int dy = Math.Max(0, Math.Max(r1.Y - r2Bottom, r2.Y - r1Bottom));
            
            return (int)Math.Sqrt(dx * dx + dy * dy);
        }

        private void DrawRectangles(Image<Rgba32> image, List<SixLabors.ImageSharp.Rectangle> rectangles)
        {
            var redColor = new Rgba32(255, 0, 0, 255); // 红色
            var lineWidth = 2; // 增加线宽，使红框更明显

            foreach (var rect in rectangles)
            {
                // 确保矩形在图像范围内
                int x1 = Math.Max(0, rect.X);
                int y1 = Math.Max(0, rect.Y);
                int x2 = Math.Min(image.Width - 1, rect.X + rect.Width - 1);
                int y2 = Math.Min(image.Height - 1, rect.Y + rect.Height - 1);
                
                // 绘制上边
                for (int x = x1; x <= x2; x++)
                {
                    for (int w = 0; w < lineWidth && y1 + w < image.Height; w++)
                    {
                        image[x, y1 + w] = redColor;
                    }
                }

                // 绘制下边
                for (int x = x1; x <= x2; x++)
                {
                    for (int w = 0; w < lineWidth && y2 - w >= 0; w++)
                    {
                        image[x, y2 - w] = redColor;
                    }
                }

                // 绘制左边
                for (int y = y1; y <= y2; y++)
                {
                    for (int w = 0; w < lineWidth && x1 + w < image.Width; w++)
                    {
                        image[x1 + w, y] = redColor;
                    }
                }

                // 绘制右边
                for (int y = y1; y <= y2; y++)
                {
                    for (int w = 0; w < lineWidth && x2 - w >= 0; w++)
                    {
                        image[x2 - w, y] = redColor;
                    }
                }
            }
        }
        
        private void DrawRectanglesWithIndexOutside_NoMeasure(
            Image<Rgba32> image,
            List<SixLabors.ImageSharp.Rectangle> rectangles)
        {
            var rectColor = Color.Red;
            var textColor = Color.DarkSlateGray;
            int lineWidth = 2;

            // 固定字体（不依赖测量）
            float fontSize = 25f;
            Font font = SystemFonts.CreateFont("Arial", fontSize, FontStyle.Bold);

            image.Mutate(ctx =>
            {
                for (int i = 0; i < rectangles.Count; i++)
                {
                    var rect = rectangles[i];

                    // 安全裁剪
                    int x = Math.Max(0, rect.X);
                    int y = Math.Max(0, rect.Y);
                    int w = Math.Min(rect.Width, image.Width - x);
                    int h = Math.Min(rect.Height, image.Height - y);

                    var safeRect = new Rectangle(x, y, w, h);

                    // ===== 1. 画矩形 =====
                    ctx.Draw(rectColor, lineWidth, safeRect);

                    // ===== 2. 编号 =====
                    string label = (i + 1).ToString();

                    PointF textPos = CalcOutsideTextPosition_NoMeasure(
                        safeRect,
                        label,
                        fontSize,
                        image.Width,
                        image.Height
                    );

                    ctx.DrawText(label, font, textColor, textPos);
                }
            });
        }
        
        private PointF CalcOutsideTextPosition_NoMeasure(
            Rectangle rect,
            string text,
            float fontSize,
            int imageWidth,
            int imageHeight)
        {
            const float padding = 4f;

            // 经验估算（对纯数字非常准）
            float approxCharWidth = fontSize * 0.6f;
            float textWidth = approxCharWidth * text.Length;
            float textHeight = fontSize;

            // 默认：左上外侧
            float x = rect.Left - textWidth - padding;
            float y = rect.Top;

            // 左侧放不下 → 放右侧
            if (x < 0)
            {
                x = rect.Right + padding;
            }

            // 上下越界保护
            if (y < 0)
            {
                y = 0;
            }
            if (y + textHeight > imageHeight)
            {
                y = imageHeight - textHeight;
            }

            // 右侧兜底
            if (x + textWidth > imageWidth)
            {
                x = Math.Max(0, rect.Left - textWidth - padding);
            }

            return new PointF(x, y);
        }

        private async Task ExportHtmlReportAsync()
        {
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
            {
                StatusMessage = "ファイルは存在していません。";
                return;
            }

            if (_topLevel == null) return;

            // 选择保存位置
            var file = await _topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "HTMLレポート保存",
                DefaultExtension = "html",
                SuggestedFileName = $"比較レポート_{DateTime.Now:yyyyMMdd_HHmmss}.html",
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("HTML ファイル")
                    {
                        Patterns = new[] { "*.html" },
                        MimeTypes = new[] { "text/html" }
                    }
                }
            });

            if (file == null || !file.Path.IsFile) return;

            IsProcessing = true;
            StatusMessage = "HTMLレポートは生成しています...";
            Progress = 0;

            try
            {
                await Task.Run(async () =>
                {
                    // ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage.License.SetNonCommercialPersonal("JiMmY");
                    
                    using (var package = new ExcelPackage(new FileInfo(FilePath)))
                    {
                        var html = new System.Text.StringBuilder();
                        html.AppendLine("<!DOCTYPE html>");
                        html.AppendLine("<html lang=\"zh-CN\">");
                        html.AppendLine("<head>");
                        html.AppendLine("<meta charset=\"UTF-8\">");
                        html.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                        html.AppendLine("<title>エビデンス比較レポート</title>");
                        html.AppendLine("<style>");
                        html.AppendLine(@"
                            body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
                            .container { max-width: 1400px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
                            h1 { color: #333; border-bottom: 3px solid #1976D2; padding-bottom: 10px; }
                            .toc { background: #f9f9f9; padding: 20px; border-radius: 5px; margin-bottom: 30px; }
                            .toc h2 { margin-top: 0; color: #1976D2; }
                            .toc-item { margin: 8px 0; padding: 8px; background: white; border-left: 3px solid #1976D2; }
                            .toc-item a { text-decoration: none; color: #333; font-weight: 500; }
                            .toc-item a:hover { color: #1976D2; }
                            .section { margin: 40px 0; padding: 20px; border: 1px solid #e0e0e0; border-radius: 5px; }
                            .section h3 { color: #1976D2; margin-top: 0; }
                            .comparison-item { margin: 20px 0; padding: 15px; background: #fafafa; border-radius: 5px; }
                            .comparison-item h4 { margin: 0 0 10px 0; color: #666; }
                            .image-container { text-align: center; margin: 10px 0; }
                            .image-container img { max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
                            .diff-info { background: #fff3cd; padding: 10px; border-radius: 4px; margin: 10px 0; }
                            .diff-info.error { background: #f8d7da; }
                            .diff-info.success { background: #d4edda; }
                            .back-to-top { position: fixed; bottom: 30px; right: 30px; background: #1976D2; color: white; border: none; padding: 15px 20px; border-radius: 50px; cursor: pointer; font-size: 16px; box-shadow: 0 4px 6px rgba(0,0,0,0.2); z-index: 1000; }
                            .back-to-top:hover { background: #1565C0; }
                        ");
                        html.AppendLine("</style>");
                        html.AppendLine("</head>");
                        html.AppendLine("<body>");
                        html.AppendLine("<div class=\"container\">");
                        html.AppendLine("<h1>画像比較レポート</h1>");
                        html.AppendLine($"<p><strong>生成時間:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");
                        html.AppendLine($"<p><strong>Excel ファイル:</strong> {FilePath}</p>");
                        
                        // 生成目录
                        html.AppendLine("<div class=\"toc\">");
                        html.AppendLine("<h2>目次</h2>");
                        
                        int totalSheets = package.Workbook.Worksheets.Count;
                        int processedSheets = 0;
                        
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            var column1Index = GetColumnIndex(Column1);
                            var column2Index = GetColumnIndex(Column2);
                        
                            // 获取使用的行数
                            int rowCount = GetActualLastRow(worksheet);

                            for (int row = 2; row <= rowCount; row++) // 从第2行开始，假设第1行是标题
                            {
                                var pic1 = GetPictureAtCell(worksheet, row, column1Index);
                                var pic2 = GetPictureAtCell(worksheet, row, column2Index);

                                if (pic1 != null && pic2 != null)
                                {
                                    html.AppendLine($"<div class=\"toc-item\"><a href=\"#sheet-{worksheet.Name}-row-{row}\">シート名「　{worksheet.Name}　」　：　行目「　{row}　」</a></div>");
                                }
                            }
                            
                            processedSheets++;
                            Dispatcher.UIThread.Post(() =>
                            {
                                Progress = (double)processedSheets / totalSheets * 20;
                                StatusMessage = $"目次は生成しています... {processedSheets}/{totalSheets}";
                            });
                        }
                        
                        html.AppendLine("</div>");
                        
                        // 生成内容
                        processedSheets = 0;
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            var column1Index = GetColumnIndex(Column1);
                            var column2Index = GetColumnIndex(Column2);
                            // 获取使用的行数
                            int rowCount = GetActualLastRow(worksheet);
                            
                            for (int row = 2; row <= rowCount; row++)
                            {
                                var pic1 = GetPictureAtCell(worksheet, row, column1Index);
                                var pic2 = GetPictureAtCell(worksheet, row, column2Index);

                                var image1Path = string.Empty;
                                var image2Path = string.Empty;
                                if (pic1 != null && pic2 != null)
                                {
                                    image1Path = SavePictureToTempFile(pic1, row, "c1");
                                    image2Path = SavePictureToTempFile(pic2, row, "c2");
                                }

                                if (image1Path == string.Empty || image2Path == string.Empty)
                                    continue;

                                // 处理图片对比
                                var comparisonResult = await ProcessImageComparisonAsync(image1Path, image2Path, row);
                                
                                html.AppendLine($"<div class=\"section\" id=\"sheet-{worksheet.Name}-row-{row}\">");
                                html.AppendLine($"<h3>シート名「　{worksheet.Name}　」　：　行目「　{row}　」</h3>");
                                html.AppendLine("<div class=\"comparison-item\">");
                                html.AppendLine($"<h4>画像パス①: {image1Path}</h4>");
                                html.AppendLine($"<h4>画像パス②: {image2Path}</h4>");
                                
                                if (comparisonResult.IsSizeMismatch)
                                {
                                    html.AppendLine($"<div class=\"diff-info error\">");
                                    html.AppendLine($"<strong>⚠ 画像のピクセルが不一致:</strong> {comparisonResult.SizeInfo}");
                                    html.AppendLine("</div>");
                                }
                                else if (comparisonResult.HasError)
                                {
                                    html.AppendLine($"<div class=\"diff-info error\">");
                                    html.AppendLine($"<strong>❌ 処理失敗:</strong> {comparisonResult.ErrorMessage}");
                                    html.AppendLine("</div>");
                                }
                                else
                                {
                                    html.AppendLine($"<div class=\"diff-info success\">");
                                    html.AppendLine($"<strong>差異数:</strong> {comparisonResult.DiffCount}");
                                    html.AppendLine("</div>");
                                    
                                    if (comparisonResult.MarkedImage1Path != null && File.Exists(comparisonResult.MarkedImage1Path))
                                    {
                                        html.AppendLine("<div class=\"image-container\">");
                                        html.AppendLine($"<p><strong>元画像① 「赤い枠でマーク」</strong></p>");
                                        html.AppendLine($"<img src=\"data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes(comparisonResult.MarkedImage1Path))}\" alt=\"原图1\">");
                                        html.AppendLine("</div>");
                                    }
                                    
                                    if (comparisonResult.MarkedImage2Path != null && File.Exists(comparisonResult.MarkedImage2Path))
                                    {
                                        html.AppendLine("<div class=\"image-container\">");
                                        html.AppendLine($"<p><strong>元画像② 「赤い枠でマーク」</strong></p>");
                                        html.AppendLine($"<img src=\"data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes(comparisonResult.MarkedImage2Path))}\" alt=\"原图2\">");
                                        html.AppendLine("</div>");
                                    }
                                    
                                    // if (comparisonResult.DiffImagePath != null && File.Exists(comparisonResult.DiffImagePath))
                                    // {
                                    //     html.AppendLine("<div class=\"image-container\">");
                                    //     html.AppendLine($"<p><strong>差異画像</strong></p>");
                                    //     html.AppendLine($"<img src=\"data:image/png;base64,{Convert.ToBase64String(File.ReadAllBytes(comparisonResult.DiffImagePath))}\" alt=\"差异图\">");
                                    //     html.AppendLine("</div>");
                                    // }
                                }
                                
                                html.AppendLine("</div>");
                                html.AppendLine("</div>");
                                
                                // 清理临时文件
                                try
                                {
                                    if (comparisonResult.MarkedImage1Path != null && File.Exists(comparisonResult.MarkedImage1Path))
                                        File.Delete(comparisonResult.MarkedImage1Path);
                                    if (comparisonResult.MarkedImage2Path != null && File.Exists(comparisonResult.MarkedImage2Path))
                                        File.Delete(comparisonResult.MarkedImage2Path);
                                    if (comparisonResult.DiffImagePath != null && File.Exists(comparisonResult.DiffImagePath))
                                        File.Delete(comparisonResult.DiffImagePath);
                                }
                                catch { }
                            }
                            
                            processedSheets++;
                            Dispatcher.UIThread.Post(() =>
                            {
                                Progress = 20 + (double)processedSheets / totalSheets * 80;
                                StatusMessage = $" シート {processedSheets}/{totalSheets} 処理中...";
                            });
                        }
                        
                        html.AppendLine("</div>");
                        html.AppendLine("<button class=\"back-to-top\" onclick=\"window.scrollTo({top: 0, behavior: 'smooth'})\">目次に戻る</button>");
                        html.AppendLine("</body>");
                        html.AppendLine("</html>");
                        
                        await File.WriteAllTextAsync(file.Path.LocalPath, html.ToString(), System.Text.Encoding.UTF8);
                    }
                });
                
                Dispatcher.UIThread.Post(() =>
                {
                    StatusMessage = "HTMLレポートは生成完了！";
                    Progress = 100;
                });
            }
            catch (Exception ex)
            {
                Dispatcher.UIThread.Post(() =>
                {
                    StatusMessage = $"生成失敗: {ex.Message}";
                    Progress = 0;
                });
            }
            finally
            {
                IsProcessing = false;
            }
        }

        private async Task<ComparisonResult> ProcessImageComparisonAsync(string image1Path, string image2Path, int rowIndex)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using var img1 = ImageSharp.Image.Load<Rgba32>(image1Path);
                    using var img2 = ImageSharp.Image.Load<Rgba32>(image2Path);
                    
                    bool isSizeMatch = img1.Width == img2.Width && img1.Height == img2.Height;
                    string sizeInfo = $"画像①: {img1.Width}x{img1.Height}, 画像②: {img2.Width}x{img2.Height}";
                    
                    if (!isSizeMatch)
                    {
                        return new ComparisonResult
                        {
                            IsSizeMismatch = true,
                            SizeInfo = sizeInfo
                        };
                    }
                    
                    using var img1Clone = img1.Clone();
                    using var img2Clone = img2.Clone();
                    
                    img1Clone.Mutate(x => x.Grayscale());
                    img2Clone.Mutate(x => x.Grayscale());
                    
                    var width = img1.Width;
                    var height = img1.Height;
                    var totalPixels = width * height;
                    var threshold = 30;
                    var differentPixels = 0;
                    var diffMap = new bool[width, height];
                    
                    for (int y = 0; y < height; y++)
                    {
                        for (int x = 0; x < width; x++)
                        {
                            var pixel1 = img1Clone[x, y];
                            var pixel2 = img2Clone[x, y];
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
                    
                    using var diffImage = img2.Clone();
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
                                    var r = Math.Min(255, Math.Abs(pixel1.R - pixel2.R) * 3);
                                    var g = Math.Min(255, Math.Abs(pixel1.G - pixel2.G) * 3);
                                    var b = Math.Min(255, Math.Abs(pixel1.B - pixel2.B) * 3);
                                    diffImage[px, y] = new Rgba32((byte)r, (byte)g, (byte)b, 255);
                                }
                            }
                        }
                    });
                    
                    var minArea = 50;
                    var mergeDistance = 10;
                    var expandPixels = 3;
                    
                    var expandedDiffMap = ExpandDiffMap(diffMap, width, height, expandPixels);
                    var differenceRegions = FindDifferenceRegions(expandedDiffMap, width, height, minArea);
                    differenceRegions = MergeNearbyRegions(differenceRegions, mergeDistance);
                    
                    using var markedImage1 = img1.Clone();
                    // DrawRectangles(markedImage1, differenceRegions);
                    DrawRectanglesWithIndexOutside_NoMeasure(markedImage1, differenceRegions);
                    
                    using var markedImage2 = img2.Clone();
                    // DrawRectangles(markedImage2, differenceRegions);
                    DrawRectanglesWithIndexOutside_NoMeasure(markedImage2, differenceRegions);
                    
                    var tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PixelCompareSuite");
                    Directory.CreateDirectory(tempDir);
                    var guid = Guid.NewGuid().ToString("N");
                    var diffImagePath = System.IO.Path.Combine(tempDir, $"diff_{rowIndex}_{guid}.png");
                    var markedImage1Path = System.IO.Path.Combine(tempDir, $"marked1_{rowIndex}_{guid}.png");
                    var markedImage2Path = System.IO.Path.Combine(tempDir, $"marked2_{rowIndex}_{guid}.png");
                    
                    diffImage.Save(diffImagePath, new PngEncoder());
                    markedImage1.Save(markedImage1Path, new PngEncoder());
                    markedImage2.Save(markedImage2Path, new PngEncoder());
                    
                    return new ComparisonResult
                    {
                        DiffCount = differenceRegions.Count,
                        DifferencePercentage = differencePercentage,
                        DiffImagePath = diffImagePath,
                        MarkedImage1Path = markedImage1Path,
                        MarkedImage2Path = markedImage2Path,
                        SizeInfo = sizeInfo
                    };
                }
                catch (Exception ex)
                {
                    return new ComparisonResult
                    {
                        HasError = true,
                        ErrorMessage = ex.Message
                    };
                }
            });
        }

        private class ComparisonResult
        {
            public double DifferencePercentage { get; set; }
            public string? DiffImagePath { get; set; }
            public string? MarkedImage1Path { get; set; }
            public string? MarkedImage2Path { get; set; }
            public bool IsSizeMismatch { get; set; }
            public string SizeInfo { get; set; } = string.Empty;
            public bool HasError { get; set; }
            public string ErrorMessage { get; set; } = string.Empty;
            public int DiffCount { get; set; }
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

        private byte[] _image1Bytes = null;
        private byte[] _image2Bytes = null;

        private int _diffCount;
        
        public int DiffCount
        {
            get => _diffCount;
            set
            {
                if (_diffCount != value)
                {
                    _diffCount = value;
                    OnPropertyChanged();
                }
            }
        }


        public byte[] Image1Bytes
        {
            get => _image1Bytes;
            set => _image1Bytes = value ?? throw new ArgumentNullException(nameof(value));
        }

        public byte[] Image2Bytes
        {
            get => _image2Bytes;
            set => _image2Bytes = value ?? throw new ArgumentNullException(nameof(value));
        }

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
                    return "ピクセルが不一致";
                if (DifferencePercentage < 0)
                    return "比較失敗";
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

