# PixelCompareSuite - Excel 图片像素对比工具

一个基于 Avalonia + MVVM 的 Excel 图片像素对比结果展示工具。

## 功能特性

- 📊 读取 Excel 文件，支持选择 Sheet 页和指定列（使用 Microsoft Excel Interop）
- 🖼️ 使用 OpenCV 进行像素级图片对比
- ⚠️ **尺寸不一致检测**：自动检测图片尺寸，尺寸不一致时在明细中标识，不进行对比
- 🔴 **差异区域红框标记**：自动检测差异区域并在原图上用红色矩形框标记
- 📄 支持分页显示对比结果列表
- 🎨 实时显示原图（带红框标记）和差异图
- 📈 显示差异百分比和可视化差异图

## 系统要求

- **.NET 9.0 SDK**
- **Windows 10+**
- **Microsoft Office** (需要安装 Excel，用于 Excel Interop)

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

1. **选择 Excel 文件**：点击"浏览..."按钮选择包含图片路径的 Excel 文件
2. **选择 Sheet 页**：从下拉框中选择要处理的 Sheet 页
3. **指定列**：输入两列名称（如 B 和 P），默认是 B 列和 P 列
4. **加载数据**：点击"加载数据"按钮读取 Excel 数据
5. **查看对比**：点击左侧列表中的对比项，右侧会自动显示：
   - 原图1
   - 原图2
   - 差异图（带差异百分比）

## Excel 文件格式

Excel 文件应包含以下格式：
- 第1行：标题行（可选）
- 从第2行开始：数据行
- 指定列（如 B 列和 P 列）包含图片文件的完整路径

示例：
| A | B | ... | P |
|---|---|-----|---|
| 标题 | 图片1路径 | ... | 图片2路径 |
| 数据1 | /path/to/image1.png | ... | /path/to/image2.png |

## 技术栈

- **Avalonia 11.0.5** - 跨平台 UI 框架
- **Microsoft Excel Interop** - Excel 文件读取（仅 Windows，需要安装 Office）
- **SixLabors.ImageSharp** - 图像处理和像素对比
- **MVVM 模式** - 视图与业务逻辑分离

## 重要说明

⚠️ **平台要求**：
- 本应用仅支持 **Windows 平台**
- 需要安装 **Microsoft Office**（包含 Excel）
- 使用 .NET 9.0 框架

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

- 图片路径必须是完整的文件系统路径
- 支持的图片格式：PNG, JPG, JPEG, BMP 等 OpenCV 支持的格式
- 差异图会临时保存在系统临时目录中
- 首次运行可能需要下载 OpenCV 运行时库

## 许可证

本项目仅供学习和参考使用。

