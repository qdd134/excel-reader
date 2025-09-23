# Excel Image Reader

一个基于TypeScript和SheetJS的Excel文件读取库，支持提取表格数据和嵌入的图片，并将图片转换为base64格式。

## 功能特性

- 📊 **完整的Excel数据解析** - 支持读取所有工作表、行、列和单元格数据
- 🖼️ **图片提取** - 自动识别和提取Excel中嵌入的图片和浮动图片
- 🔄 **Base64转换** - 将图片转换为base64格式，便于Web使用
- 📝 **类型安全** - 完整的TypeScript类型定义
- 🎯 **灵活配置** - 支持多种解析选项和自定义配置
- 📈 **详细报告** - 生成JSON、CSV和HTML格式的详细报告
- 🔢 **多图片支持** - 支持一行中包含多张图片的复杂场景
- 📊 **图片统计** - 提供详细的图片分布和统计信息
- 🚀 **零配置运行** - 提供JavaScript版本，无需编译即可使用
- 🏗️ **模块化架构** - 组件化设计，易于扩展和维护
- 🎨 **浮动图片支持** - 完整支持Excel中的浮动图片解析和定位

## 仓库地址

- GitHub: `https://github.com/qdd134/excel-reader.git`

## 安装

支持两种方式：通过 GitHub 直接安装，或克隆/下载源码本地使用。

### 方式一：作为依赖从 GitHub 安装（推荐）

```bash
# 使用 HTTPS（推荐）
npm i git+https://github.com/qdd134/excel-reader.git#v1.2.0
# 或使用 github 简写（yarn/pnpm支持）
yarn add github:qdd134/excel-reader#v1.2.0
```

安装后在你的项目中直接使用：

```typescript
import { ExcelImageReader } from 'excel-reader';

const reader = new ExcelImageReader();
const result = await reader.parseFile('path/to/your.xlsx', { includeImages: true });
```

> 提示：仓库包含 `"prepare": "npm run build"`，用 Git 方式安装时会自动编译生成 `dist/`。

### 方式二：克隆仓库本地使用

```bash
git clone https://github.com/qdd134/excel-reader.git
cd excel-reader
npm install
npm run build
```

## 快速开始

### 基本使用

```typescript
// 若通过 GitHub 安装依赖，请从包名导入
import { ExcelImageReader } from 'excel-reader';
// 若直接在本仓库内开发，请从 src 导入：
// import { ExcelImageReader } from './src/index';

const reader = new ExcelImageReader();

// 从文件路径读取
const result = await reader.parseFile('path/to/your/file.xlsx', {
  includeImages: true,
  includeEmptyRows: false,
  includeEmptyColumns: false
});

// 从Buffer读取（Node 环境）
import * as fs from 'fs';
const buffer = fs.readFileSync('path/to/your/file.xlsx');
const result2 = await reader.parseBuffer(buffer, { includeImages: true });

console.log(`发现 ${result.worksheets.length} 个工作表`);
console.log(`发现 ${result.images.size} 张图片`);
```

### 处理结果数据

```typescript
for (const worksheet of result.worksheets) {
  console.log(`工作表: ${worksheet.name}`);
  console.log(`总图片数: ${worksheet.totalImages}`);
  console.log(`包含图片的行数: ${worksheet.rowsWithImages}`);
  for (const row of worksheet.rows) {
    if (row.imageCount > 0) {
      console.log(`  📷 该行包含 ${row.imageCount} 张图片，位置: ${row.imageCells.join(', ')}`);
    }
    for (const cell of row.cells) {
      if (cell.image) {
        console.log(`  ${cell.ref}: [图片] ${cell.image.description}`);
      } else {
        console.log(`  ${cell.ref}: ${cell.value}`);
      }
    }
  }
}
```

### HTML 预览（数据 + 图片一起展示）

```bash
# 在仓库根目录
npm run build && node dist/example/example.js
# 打开生成的预览文件
open example/output/preview.html
```

> 预览文件固定输出到 `example/output/preview.html`，`test.xlsx` 请放在 `example/test.xlsx`。

### 多图片处理

```typescript
const reader = new ExcelImageReader();
const result = await reader.parseFile('file.xlsx', { includeImages: true });

for (const worksheet of result.worksheets) {
  const multiImageRows = worksheet.rows.filter(row => row.imageCount > 1);
  console.log(`多图片行数: ${multiImageRows.length}`);
}
```

### 高级多图片处理

```typescript
import { MultiImageProcessor } from './example/multi-image-example';

const processor = new MultiImageProcessor();
await processor.processMultiImageExcel('file.xlsx');
await processor.extractMultiImageRows('file.xlsx');
```

### 浮动图片处理

```typescript
import { ExcelImageReader } from 'excel-reader';

const reader = new ExcelImageReader();
const result = await reader.parseFile('file.xlsx', { includeImages: true });

// 处理浮动图片
for (const worksheet of result.worksheets) {
  for (const row of worksheet.rows) {
    if (row.floatingImages && row.floatingImages.length > 0) {
      console.log(`行 ${row.rowNumber} 包含 ${row.floatingImages.length} 张浮动图片`);
      
      row.floatingImages.forEach(image => {
        console.log(`  - 图片ID: ${image.id}`);
        console.log(`  - 位置: (${image.position.x}, ${image.position.y})`);
        console.log(`  - 尺寸: ${image.position.width} x ${image.position.height}`);
      });
    }
  }
}
```

### 组件独立使用

```typescript
import { DrawingParser, RelationshipParser, ImageExtractor } from 'excel-reader';

// 独立使用DrawingParser
const drawingParser = new DrawingParser();
const floatingImages = drawingParser.parseDrawingXml(drawingXml);

// 独立使用RelationshipParser
const relationshipParser = new RelationshipParser();
const relationships = relationshipParser.parseRelationships(relsXml);

// 独立使用ImageExtractor
const imageExtractor = new ImageExtractor();
imageExtractor.setZip(zipInstance);
const imageData = await imageExtractor.extractImageData('xl/media/image1.jpeg');
```

## 作为依赖使用（无 npm 发布）

- 通过 GitHub 直接引用（见上文“方式一”）。
- 或生成 tar 包：
  ```bash
  npm run build
  npm pack
  # 生成 excel-reader-<version>.tgz 进行安装
  npm i /absolute/path/to/excel-image-reader-<version>.tgz
  ```

## API 文档

### 主要类

#### ExcelImageReader

主要的Excel解析器类，负责协调各个组件完成Excel文件的解析。

#### 方法

##### `parseFile(filePath: string, options?: ParseOptions): Promise<ExcelParseResult>`

从文件路径解析Excel文件。

**参数:**
- `filePath` - Excel文件路径
- `options` - 解析选项（可选）

**返回:** Promise<ExcelParseResult>

##### `parseBuffer(buffer: Buffer, options?: ParseOptions): Promise<ExcelParseResult>`

从Buffer解析Excel文件。

**参数:**
- `buffer` - Excel文件Buffer
- `options` - 解析选项（可选）

**返回:** Promise<ExcelParseResult>

### 组件类

#### DrawingParser

专门解析Excel中的drawing.xml文件，提取浮动图片信息。

```typescript
import { DrawingParser } from 'excel-reader';

const parser = new DrawingParser();
const floatingImages = parser.parseDrawingXml(drawingXml);
```

#### RelationshipParser

解析Excel中的各种关系文件（.rels），处理文件间的依赖关系。

```typescript
import { RelationshipParser } from 'excel-reader';

const parser = new RelationshipParser();
const relationships = parser.parseRelationships(relsXml);
```

#### ImageExtractor

负责从Excel文件中提取图片的二进制数据并转换为base64格式。

```typescript
import { ImageExtractor } from 'excel-reader';

const extractor = new ImageExtractor();
const imageData = await extractor.extractImageData('xl/media/image1.jpeg');
```

#### FloatingImageManager

管理浮动图片的解析、绑定和归类，协调其他组件完成浮动图片处理。

```typescript
import { FloatingImageManager } from 'excel-reader';

const manager = new FloatingImageManager();
await manager.parseFloatingImages(result);
```

### 类型定义

#### ParseOptions

```typescript
interface ParseOptions {
  includeImages?: boolean;
  imageQuality?: number;
  includeEmptyRows?: boolean;
  includeEmptyColumns?: boolean;
}
```

#### ExcelParseResult

```typescript
interface ExcelParseResult {
  worksheets: WorksheetData[];
  images: Map<string, CellImage>;
  errors: string[];
}
```

#### CellImage

```typescript
interface CellImage {
  id: string;
  description: string;
  base64: string;
  mimeType: string;
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  relationshipId: string;
}
```

#### RowData

```typescript
interface RowData {
  rowNumber: number;
  height?: number;
  customHeight?: boolean;
  cells: CellData[];
  imageCount: number;
  imageCells: string[];
  floatingImageIds?: string[];
  floatingImages?: CellImage[];
}
```

#### WorksheetData

```typescript
interface WorksheetData {
  name: string;
  dimension: {
    start: string;
    end: string;
  };
  rows: RowData[];
  columns: {
    min: number;
    max: number;
    width: number;
    customWidth: boolean;
  }[];
  totalImages: number;
  rowsWithImages: number;
}
```

## 示例

- `example/example.ts` 基础示例
- `example/advanced-example.ts` 高级功能（报告、CSV、HTML）
- `example/multi-image-example.ts` 多图片处理

## 运行示例

```bash
npm install
npm run build
npm run test         # JS 简化示例
npm run example      # 构建后运行
node example/simple-example.js

#可以使用npx临时运行示例
cd example
npx ts-node example.ts
```

## 项目结构

```
excel-reader/
├── src/
│   ├── types.ts                   # 类型定义
│   ├── ExcelImageReader.ts        # 主解析器类
│   ├── DrawingParser.ts           # Drawing XML解析器
│   ├── RelationshipParser.ts      # 关系文件解析器
│   ├── ImageExtractor.ts          # 图片数据提取器
│   ├── FloatingImageManager.ts    # 浮动图片管理器
│   └── index.ts                   # 导出文件
├── example/
│   ├── simple-example.js
│   ├── example.ts
│   ├── advanced-example.ts
│   └── multi-image-example.ts
├── dist/                          # 编译输出目录
├── package.json
├── tsconfig.json
├── REFACTORING_SUMMARY.md         # 重构总结文档
└── README.md
```

## 支持的Excel功能

- ✅ 基本单元格数据（文本、数字、公式）
- ✅ 图片嵌入（JPEG、PNG、GIF、BMP、WebP）
- ✅ 浮动图片（Floating Images）
- ✅ 工作表结构
- ✅ 行高和列宽设置
- ✅ 单元格样式
- ✅ 公式计算
- ✅ 图片位置和尺寸信息
- ✅ 图片锚点定位
- ✅ 多图片场景支持

## 架构设计

### 组件化架构

本项目采用组件化设计，将复杂的Excel解析逻辑拆分为多个独立的组件：

```
ExcelImageReader (主协调器)
├── FloatingImageManager (浮动图片管理器)
│   ├── DrawingParser (Drawing XML解析器)
│   ├── RelationshipParser (关系文件解析器)
│   └── ImageExtractor (图片数据提取器)
├── RelationshipParser (嵌入式图片关系解析)
└── ImageExtractor (嵌入式图片数据提取)
```

### 组件职责

- **ExcelImageReader**: 主解析器，协调各个组件完成Excel文件解析
- **DrawingParser**: 专门解析drawing.xml文件，提取浮动图片信息
- **RelationshipParser**: 解析各种关系文件（.rels），处理文件间依赖
- **ImageExtractor**: 提取图片二进制数据并转换为base64格式
- **FloatingImageManager**: 管理浮动图片的完整生命周期

### 设计优势

- 🎯 **单一职责**: 每个组件都有明确的职责边界
- 🧪 **可测试性**: 组件可以独立进行单元测试
- 🔧 **可扩展性**: 新功能可以轻松添加到相应组件
- 📦 **可重用性**: 组件可以在其他项目中重用
- 🛠️ **可维护性**: 代码结构清晰，易于理解和修改

## 技术栈

- **TypeScript** - 类型安全的JavaScript超集
- **SheetJS (xlsx)** - Excel文件解析核心库
- **JSZip** - ZIP文件处理库
- **fast-xml-parser** - XML解析库
- **Node.js** - 运行时环境

## 注意事项

1. 确保Excel文件是有效的XLSX格式
2. 图片提取需要 `xl/cellimages.xml` 与关联图片存在
3. 浮动图片需要 `xl/drawings/` 目录下的drawing文件
4. Base64体积较大时注意内存与输出体积
5. 复杂特性可能不完全覆盖
6. 快速预览：`example/test.xlsx` + 生成 `example/output/preview.html`
7. 组件化架构支持独立使用各个解析器

## 贡献

欢迎提交Issue和Pull Request来改进这个库！

## 开源协议

MIT License

## 更新日志

### v1.2.0
- 🏗️ **架构重构** - 采用组件化设计，提高代码可维护性
- 🎨 **浮动图片支持** - 完整支持Excel中的浮动图片解析和定位
- 🔧 **模块化组件** - 新增DrawingParser、RelationshipParser、ImageExtractor、FloatingImageManager组件
- 📈 **性能优化** - 优化图片解析流程，提高处理效率
- 🧪 **代码质量** - 提高代码可测试性和可扩展性
- 📚 **文档完善** - 更新API文档和架构说明

### v1.1.0
- 新增多图片支持
- 增强数据结构，添加图片统计信息
- 提供JavaScript示例
- 优化错误处理与类型兼容
- 添加图片分布分析与多图片行提取

### v1.0.0
- 初始版本发布
- 支持基本的Excel数据解析与图片提取
- 完整的TypeScript类型定义
- 提供基础/高级示例
