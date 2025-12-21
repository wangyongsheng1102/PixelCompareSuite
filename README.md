# PixelCompareSuite - Excel 图片像素对比工具

一个基于 Avalonia + MVVM 的 Excel 图片像素对比结果展示工具，用于批量对比 Excel 文件中指定列的图片，并可视化显示像素级差异。

## 功能特性

- 📊 **Excel 文件读取**：支持选择 Sheet 页和指定列（使用 Microsoft Excel Interop）
- 🖼️ **像素级对比**：使用 SixLabors.ImageSharp 进行高精度像素级图片对比
- ⚠️ **尺寸不一致检测**：自动检测图片尺寸，尺寸不一致时在明细中标识，不进行对比
- 🔴 **差异区域红框标记**：自动检测差异区域（连通域分析）并在原图上用红色矩形框标记
- 📄 **分页显示**：支持分页显示对比结果列表，提高大量数据时的性能
- 🎨 **实时预览**：实时显示原图（带红框标记）和差异图
- 📈 **差异统计**：显示差异百分比和可视化差异图
- ⚡ **异步处理**：使用异步加载，不阻塞 UI 界面
- 🎯 **智能阈值**：可配置的像素差异阈值（默认 30），过滤微小差异

## 系统要求

### 必需组件
- **.NET 9.0 SDK** 或更高版本
- **Windows 10/11** (64位)
- **Microsoft Office** (需要安装 Excel 2016 或更高版本，用于 Excel Interop)

### 推荐配置
- **内存**：8GB RAM 或更多（处理大量图片时）
- **磁盘空间**：至少 500MB 可用空间（用于临时文件）
- **处理器**：多核处理器（支持并行处理）

## 运行项目

### 1. 恢复依赖包

```bash
dotnet restore
```

### 2. 构建项目

```bash
dotnet build
```

### 3. 运行项目

```bash
dotnet run
```

或者直接运行可执行文件：

```bash
dotnet run --project PixelCompareSuite.csproj
```

## 使用说明

### 基本操作流程

1. **选择 Excel 文件**
   - 点击"浏览..."按钮选择包含图片路径的 Excel 文件
   - 支持 `.xlsx` 和 `.xls` 格式

2. **选择 Sheet 页**
   - 从下拉框中选择要处理的 Sheet 页
   - 程序会自动读取所有可用的 Sheet 名称

3. **指定列**
   - 输入两列名称（如 `B` 和 `P`），默认是 B 列和 P 列
   - 列名不区分大小写
   - 第一列对应"原图1"，第二列对应"原图2"

4. **加载数据**
   - 点击"加载数据"按钮读取 Excel 数据
   - 程序会从第2行开始读取（第1行视为标题行）
   - 底部进度条会显示加载进度

5. **查看对比结果**
   - 点击左侧列表中的对比项，右侧会自动显示：
     - **原图1**：第一列对应的图片（带差异区域红框标记）
     - **原图2**：第二列对应的图片（带差异区域红框标记）
     - **差异图**：高亮显示像素差异的区域（带差异百分比）
   - 使用分页控件可以浏览更多对比项

### 差异图说明

- **差异百分比**：显示两张图片的像素差异百分比
- **红框标记**：在原图上用红色矩形框标记差异区域（最小区域面积：50 像素）
- **尺寸不一致**：如果两张图片尺寸不同，会显示"尺寸不一致"提示，并显示具体尺寸信息

## Excel 文件格式

### 文件结构要求

Excel 文件应包含以下格式：
- **第1行**：标题行（可选，程序会自动跳过）
- **从第2行开始**：数据行，每行代表一组需要对比的图片
- **指定列**（如 B 列和 P 列）包含图片文件的**完整绝对路径**

### 路径格式

- 支持 Windows 路径格式：`C:\Users\Username\Pictures\image1.png`
- 支持网络路径：`\\server\share\image.png`
- 路径中不能包含 Excel 公式，必须是纯文本路径

### 示例

| A | B | ... | P |
|---|---|-----|---|
| 序号 | 图片1路径 | ... | 图片2路径 |
| 1 | C:\Images\test1.png | ... | C:\Images\test2.png |
| 2 | C:\Images\sample1.jpg | ... | C:\Images\sample2.jpg |

### 注意事项

- 如果某行的图片路径为空或无效，该行会被跳过
- 图片路径必须是有效的文件系统路径
- 程序会验证文件是否存在，不存在的文件会标记为错误

## 技术栈

### 核心框架
- **.NET 9.0** - 应用程序运行时
- **Avalonia 11.0.5** - 现代化 UI 框架（Windows 平台）
- **MVVM 模式** - 视图与业务逻辑分离，使用数据绑定

### 主要依赖库
- **Microsoft.Office.Interop.Excel 16.0.0** - Excel 文件读取（仅 Windows，需要安装 Office）
- **SixLabors.ImageSharp 3.1.5** - 高性能图像处理和像素对比
- **SixLabors.ImageSharp.Drawing 2.1.0** - 图像绘制功能（用于红框标记）
- **Avalonia.ReactiveUI 11.0.5** - 响应式 UI 编程

### 技术特点
- **异步编程**：使用 `async/await` 模式，确保 UI 响应流畅
- **COM 互操作**：正确管理 Excel Interop COM 对象，防止内存泄漏
- **像素级对比算法**：
  - 灰度转换后对比，减少颜色差异干扰
  - 可配置阈值（默认 30）过滤微小差异
  - 连通域分析（BFS）检测差异区域
  - 最小区域过滤（50 像素）减少噪声

## 重要说明

### ⚠️ 平台限制
- **本应用仅支持 Windows 平台**（Windows 10/11 64位）
- 由于使用 Microsoft Excel Interop，无法在 Linux 或 macOS 上运行
- 需要安装 **Microsoft Office**（包含 Excel 2016 或更高版本）

### 🔒 权限要求
- 应用程序需要访问 Excel 文件和图片文件的读取权限
- 临时文件会保存在系统临时目录（`%TEMP%\PixelCompareSuite`）

### 💾 内存管理
- 程序会自动释放 Excel COM 对象，避免内存泄漏
- 差异图会临时保存在磁盘上，处理完成后可手动清理
- 建议处理大量图片时关闭其他占用内存的程序

## 项目结构

```
PixelCompareSuite/
├── App.axaml                      # 应用程序定义
├── App.axaml.cs                   # 应用程序代码
├── Program.cs                     # 程序入口
├── Views/                         # 视图目录
│   ├── CompareResultView.axaml    # 主视图
│   └── CompareResultView.axaml.cs # 视图代码
├── ViewModels/                    # 视图模型目录
│   └── CompareResultViewModel.cs  # 视图模型
├── Converters/                    # 值转换器目录
│   ├── ObjectToBoolConverter.cs  # 布尔值转换器
│   └── PathToBitmapConverter.cs   # 图片路径转换器
├── PixelCompareSuite.csproj       # 项目文件
├── app.manifest                   # 应用程序清单
├── Roots.xml                      # Trimmer 配置
└── Assets/                        # 资源文件目录
```

## 注意事项

### 图片要求
- **路径格式**：必须是完整的绝对路径（如 `C:\Images\test.png`）
- **支持格式**：PNG, JPG, JPEG, BMP, GIF, WebP 等 ImageSharp 支持的所有格式
- **文件大小**：建议单张图片不超过 50MB，过大的图片可能影响性能
- **图片尺寸**：如果两张图片尺寸不一致，程序会标记但不会进行对比

### 性能优化
- **分页显示**：默认每页显示 20 项，可通过分页控件浏览
- **异步加载**：图片对比在后台线程进行，不会阻塞 UI
- **临时文件**：差异图保存在 `%TEMP%\PixelCompareSuite`，可定期清理

### 常见问题

**Q: 为什么程序无法打开 Excel 文件？**
- A: 确保已安装 Microsoft Office（包含 Excel），并且 Excel 文件未被其他程序占用

**Q: 图片对比结果不准确？**
- A: 检查图片路径是否正确，确保图片文件存在且可读。如果图片尺寸不一致，程序会标记但不会对比

**Q: 程序运行缓慢？**
- A: 处理大量图片时，建议分批处理。确保有足够的内存和磁盘空间

**Q: 差异图显示不正确？**
- A: 检查临时目录是否有写入权限，确保 `%TEMP%\PixelCompareSuite` 目录可写

**Q: COM 对象错误？**
- A: 确保 Excel 已正确安装，并且没有其他程序正在使用 Excel Interop

## 开发说明

### 构建项目

```bash
# 恢复 NuGet 包
dotnet restore

# 编译项目（Debug 模式）
dotnet build

# 编译项目（Release 模式）
dotnet build -c Release

# 运行项目
dotnet run
```

### 发布项目

```bash
# 发布为单文件可执行程序
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# 发布后的文件在 bin/Release/net9.0/win-x64/publish/ 目录
```

### 代码结构说明

- **ViewModels/CompareResultViewModel.cs**：核心业务逻辑，包含 Excel 读取、图片对比算法
- **Views/CompareResultView.axaml**：UI 布局定义
- **Converters/**：值转换器，用于数据绑定转换
- **app.manifest**：应用程序清单，包含 Windows 权限配置

### 关键算法

1. **像素对比算法**：
   - 将图片转换为灰度图
   - 逐像素比较，计算差异
   - 使用阈值过滤微小差异

2. **连通域分析**：
   - 使用 BFS（广度优先搜索）找到差异区域
   - 过滤小于最小面积的区域
   - 计算每个区域的边界框

3. **红框绘制**：
   - 在原始图片上绘制红色矩形框
   - 使用像素级操作确保精确绘制

## 许可证

本项目仅供学习和参考使用。

## 更新日志

### v1.0.0
- ✅ 初始版本发布
- ✅ 支持 Excel 文件读取和图片对比
- ✅ 实现像素级差异检测和红框标记
- ✅ 支持分页显示和异步处理
- ✅ 尺寸不一致检测功能

