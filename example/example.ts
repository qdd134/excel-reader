// @ts-nocheck
import { ExcelImageReader, ExcelParseResult, CellData } from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import dotenv from 'dotenv';
dotenv.config({ path: '.env' });

/**
 * 解析 example/test.xlsx 的真实路径（兼容编译后 dist/example 运行）
 */
function resolveExampleXlsx(): string {
  // 优先使用环境变量指定的路径
  const envPath = process.env.EXCEL_READER_XLSX || process.env.EXCEL_XLSX || process.env.XLSX_PATH;
  if (envPath && fs.existsSync(envPath)) {
    return envPath;
  }

  const candidates = [
    path.join(__dirname, 'test.xlsx'),
    path.resolve(__dirname, '../../example/test.xlsx'),
    path.resolve(__dirname, '../test.xlsx')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return candidates[0];
}

/**
 * 解析根目录与是否全量扫描
 */
function resolveRootAndMode(): { rootDir: string | null; parseAll: boolean } {
  const rootEnv = process.env.EXCEL_READER_ROOT || process.env.EXCEL_ROOT || process.env.XLSX_ROOT || '';
  const parseAllEnv = process.env.EXCEL_READER_PARSE_ALL || process.env.EXCEL_PARSE_ALL || process.env.PARSE_ALL || '';
  const parseAll = /^(1|true|yes)$/i.test(parseAllEnv.trim());
  const rootDir = rootEnv ? path.resolve(rootEnv) : null;
  return { rootDir, parseAll };
}

/**
 * 扫描根目录下的所有xlsx文件（非递归）
 */
function listXlsxFiles(dir: string): string[] {
  if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) return [];
  return fs.readdirSync(dir)
    .filter(n => n.toLowerCase().endsWith('.xlsx'))
    .map(n => path.join(dir, n));
}

/**
 * 获取输出目录（固定写到源代码 example/output 目录）
 */
function resolveOutputDir(): string {
  const envOut = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR;
  const out = envOut ? path.resolve(envOut) : path.resolve(__dirname, '../../example/output');
  if (!fs.existsSync(out)) fs.mkdirSync(out, { recursive: true });
  return out;
}

/**
 * 示例：读取Excel文件并提取图片数据
 */
async function main() {
  const reader = new ExcelImageReader();

  const { rootDir, parseAll } = resolveRootAndMode();
  const excelFilePath = resolveExampleXlsx();
  const targets = parseAll && rootDir ? listXlsxFiles(rootDir) : [excelFilePath];

  try {
    console.log('开始解析Excel文件...');
    console.log(`模式: ${parseAll ? '全量解析目录' : '单文件解析'}`);
    if (parseAll && rootDir) console.log(`根目录: ${rootDir}`);

    const jsonFiles: string[] = [];
    for (const file of targets) {
      console.log(`\n=== 解析: ${file} ===`);
      const result: ExcelParseResult = await reader.parseFile(file, {
        includeImages: true,
        includeEmptyRows: false,
        includeEmptyColumns: true
      });

      console.log(`解析完成！发现 ${result.worksheets.length} 个工作表`);
      console.log(`发现 ${result.images.size} 张图片`);

      if (result.errors.length > 0) {
        console.log('解析过程中的错误：');
        result.errors.forEach(error => console.log(`- ${error}`));
      }

      // 保存解析结果为 JSON（供多文件 HTML 预览按需加载）
      const jf = await saveResultJsonToFile(result, file);
      jsonFiles.push(jf);

      // （已移除）导出到xlsx

      //遍历所有工作表（日志）
      for (const worksheet of result.worksheets) {
        console.log(`\n=== 工作表: ${worksheet.name} ===`);
        console.log(`数据范围: ${worksheet.dimension.start}:${worksheet.dimension.end}`);
        console.log(`行数: ${worksheet.rows.length}`);
        console.log(`图片总数: ${worksheet.totalImages}`);
        console.log(`包含图片的行数: ${worksheet.rowsWithImages}`);

        // 遍历所有行
        // for (const row of worksheet.rows) {
        //   console.log(`\n--- 第 ${row.rowNumber} 行 ---`);
        //   if (row.height) {
        //     console.log(`行高: ${row.height}pt`);
        //   }

        //   // 显示该行的图片统计
        //   if (row.imageCount > 0) {
        //     console.log(`📷 该行包含 ${row.imageCount} 张图片，位置: ${row.imageCells.join(', ')}`);
        //   }

        //   // 遍历所有单元格
        //   for (const cell of row.cells) {
        //     console.log(`单元格 ${cell.ref}: ${cell.value} (类型: ${cell.type})`);

        //     // 如果是图片单元格，显示图片信息
        //     if (cell.image) {
        //       console.log(`  └─ 图片ID: ${cell.image.id}`);
        //       console.log(`  └─ 图片描述: ${cell.image.description}`);
        //       console.log(`  └─ 图片尺寸: ${cell.image.position.width}x${cell.image.position.height}`);
        //       console.log(`  └─ 图片位置: (${cell.image.position.x}, ${cell.image.position.y})`);
        //       console.log(`  └─ 图片MIME类型: ${cell.image.mimeType}`);
        //       console.log(`  └─ Base64长度: ${cell.image.base64.length} 字符`);

        //       // 保存图片到文件（可选）
        //       await saveImageToFile(cell.image, worksheet.name, cell.ref);
        //     }
        //   }
        // }
      }

      // 精简日志，避免大文件卡顿
      console.log(`图片统计: ${result.images.size}`);
    }

    // 生成一个多文件 HTML 索引，按需加载 JSON，提高大表打开速度
    await generateMultiFileHtmlIndex(jsonFiles);

  } catch (error) {
    console.error('解析失败:', error);
  }
}

/**
 * 保存图片到文件
 */
async function saveImageToFile(image: any, worksheetName: string, cellRef: string): Promise<void> {
  try {
    const outputDir = resolveOutputDir();

    const fileName = `${worksheetName}_${cellRef}_${image.id}.${getFileExtension(image.mimeType)}`;
    const filePath = path.join(outputDir, fileName);

    // 提取base64数据（去掉data:image/xxx;base64,前缀）
    const base64Data = image.base64.split(',')[1];
    const buffer = Buffer.from(base64Data, 'base64');

    fs.writeFileSync(filePath, buffer);
    console.log(`  └─ 图片已保存到: ${filePath}`);
  } catch (error) {
    console.error(`保存图片失败: ${error}`);
  }
}

// （已移除）旧版内联 HTML 预览模板

/**
 * 根据MIME类型获取文件扩展名
 */
function getFileExtension(mimeType: string): string {
  switch (mimeType) {
    case 'image/jpeg':
      return 'jpg';
    case 'image/png':
      return 'png';
    case 'image/gif':
      return 'gif';
    case 'image/bmp':
      return 'bmp';
    case 'image/webp':
      return 'webp';
    default:
      return 'jpg';
  }
}

/**
 * 示例：从Buffer读取Excel文件
 */
async function readFromBuffer() {
  const reader = new ExcelImageReader();

  try {
    // 读取文件为Buffer（兼容运行位置）
    const filePath = resolveExampleXlsx();
    const buffer = fs.readFileSync(filePath);

    console.log('从Buffer读取Excel文件...');
    const result = await reader.parseBuffer(buffer, {
      includeImages: true
    });

    console.log('Buffer解析完成！');
    console.log(`工作表数量: ${result.worksheets.length}`);
    console.log(`图片数量: ${result.images.size}`);

  } catch (error) {
    console.error('Buffer解析失败:', error);
  }
}

// 运行示例
if (require.main === module) {
  main().then(() => {
    console.log('\n示例执行完成！');
  }).catch(error => {
    console.error('示例执行失败:', error);
  });
}

export { main, readFromBuffer };

/**
 * 保存解析结果为 JSON（将 Map 转为普通对象）
 */
async function saveResultJson(result: ExcelParseResult): Promise<void> {
  const outDir = resolveOutputDir();
  const outPath = path.join(outDir, 'result.json');
  const imagesObj: Record<string, any> = {};
  result.images.forEach((v, k) => { imagesObj[k] = v; });
  const serializable = {
    worksheets: result.worksheets,
    images: imagesObj,
    errors: result.errors
  };
  fs.writeFileSync(outPath, JSON.stringify(serializable, null, 2), 'utf8');
  console.log(`结果 JSON 已保存: ${outPath}`);
}

// 将结果保存为 output/<base>/<base>.json，并将图片写入 output/<base>/images/
// JSON 中不再内嵌 base64，改为提供相对 url，显著减小体积；返回相对 JSON 路径
async function saveResultJsonToFile(result: ExcelParseResult, sourcePath: string): Promise<string> {
  const outDir = resolveOutputDir();
  const base = path.basename(sourcePath, path.extname(sourcePath));
  const baseDir = path.join(outDir, base);
  const imagesDir = path.join(baseDir, 'images');
  if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });
  if (!fs.existsSync(imagesDir)) fs.mkdirSync(imagesDir, { recursive: true });

  const outFile = `${base}/${base}.json`;
  const outPath = path.join(outDir, outFile);

  // 写出图片文件，并构造精简的 images 映射（移除 base64）
  const imagesObj: Record<string, any> = {};
  result.images.forEach((img, id) => {
    try {
      const ext = getFileExtension(img.mimeType);
      const filename = `${id}.${ext}`;
      const filePath = path.join(imagesDir, filename);
      // 写文件
      const base64Data = (img.base64 || '').split(',')[1] || '';
      if (base64Data) {
        const buffer = Buffer.from(base64Data, 'base64');
        fs.writeFileSync(filePath, buffer);
      }
      // 在 JSON 中仅保留元信息与相对 url（相对于 output/preview.html）
      imagesObj[id] = {
        id: img.id,
        description: img.description,
        mimeType: img.mimeType,
        position: img.position,
        relationshipId: (img as any).relationshipId,
        url: `${base}/images/${filename}`
      };
    } catch {}
  });
  const serializable = {
    worksheets: result.worksheets,
    images: imagesObj,
    errors: result.errors
  };
  fs.writeFileSync(outPath, JSON.stringify(serializable), 'utf8');
  console.log(`结果 JSON 已保存: ${outPath}`);
  return outFile;
}

// 生成多文件索引 HTML：点击展开、分页渲染、图片 lazy load，显著优化大表打开速度
async function generateMultiFileHtmlIndex(jsonFiles: string[]): Promise<void> {
  const outDir = resolveOutputDir();
  const outPath = path.join(outDir, 'preview.html');
  const tplPath = path.resolve(__dirname, './templates/preview.html');
  let tpl = fs.readFileSync(tplPath, 'utf8');
  tpl = tpl.replace('/*__FILES__*/[]', JSON.stringify(jsonFiles));
  fs.writeFileSync(outPath, tpl, 'utf8');
  console.log(`多文件 HTML 预览已生成: ${outPath}`);
}

// （删除）导出解析结果为 xlsx 的逻辑

function safeSheetName(name: string): string {
  const n = name.replace(/[\\\/:\?\*\[\]]/g, '_');
  return n.length > 31 ? n.slice(0, 31) : n;
}

