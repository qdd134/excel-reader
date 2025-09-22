// @ts-nocheck
import { ExcelImageReader, ExcelParseResult, CellData } from '../src/index';
import * as fs from 'fs';
import * as path from 'path';

/**
 * 解析 example/test.xlsx 的真实路径（兼容编译后 dist/example 运行）
 */
function resolveExampleXlsx(): string {
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
 * 获取输出目录（固定写到源代码 example/output 目录）
 */
function resolveOutputDir(): string {
  const out = path.resolve(__dirname, '../../example/output');
  if (!fs.existsSync(out)) fs.mkdirSync(out, { recursive: true });
  return out;
}

/**
 * 示例：读取Excel文件并提取图片数据
 */
async function main() {
  const reader = new ExcelImageReader();
  
  // 兼容 src 和 dist 下运行
  const excelFilePath = resolveExampleXlsx();
  
  try {
    console.log('开始解析Excel文件...');
    console.log(`使用文件: ${excelFilePath}`);
    
    // 解析Excel文件
    const result: ExcelParseResult = await reader.parseFile(excelFilePath, {
      includeImages: true,
      includeEmptyRows: false,
      includeEmptyColumns: false
    });

    console.log(`解析完成！发现 ${result.worksheets.length} 个工作表`);
    console.log(`发现 ${result.images.size} 张图片`);
    
    if (result.errors.length > 0) {
      console.log('解析过程中的错误：');
      result.errors.forEach(error => console.log(`- ${error}`));
    }

    // 生成HTML预览（数据 + 图片）
    await generateHtmlPreview(result);

    // 遍历所有工作表
    for (const worksheet of result.worksheets) {
      console.log(`\n=== 工作表: ${worksheet.name} ===`);
      console.log(`数据范围: ${worksheet.dimension.start}:${worksheet.dimension.end}`);
      console.log(`行数: ${worksheet.rows.length}`);
      console.log(`图片总数: ${worksheet.totalImages}`);
      console.log(`包含图片的行数: ${worksheet.rowsWithImages}`);
      
      // 遍历所有行
      for (const row of worksheet.rows) {
        console.log(`\n--- 第 ${row.rowNumber} 行 ---`);
        if (row.height) {
          console.log(`行高: ${row.height}pt`);
        }
        
        // 显示该行的图片统计
        if (row.imageCount > 0) {
          console.log(`📷 该行包含 ${row.imageCount} 张图片，位置: ${row.imageCells.join(', ')}`);
        }
        
        // 遍历所有单元格
        for (const cell of row.cells) {
          console.log(`单元格 ${cell.ref}: ${cell.value} (类型: ${cell.type})`);
          
          // 如果是图片单元格，显示图片信息
          if (cell.image) {
            console.log(`  └─ 图片ID: ${cell.image.id}`);
            console.log(`  └─ 图片描述: ${cell.image.description}`);
            console.log(`  └─ 图片尺寸: ${cell.image.position.width}x${cell.image.position.height}`);
            console.log(`  └─ 图片位置: (${cell.image.position.x}, ${cell.image.position.y})`);
            console.log(`  └─ 图片MIME类型: ${cell.image.mimeType}`);
            console.log(`  └─ Base64长度: ${cell.image.base64.length} 字符`);
            
            // 保存图片到文件（可选）
            await saveImageToFile(cell.image, worksheet.name, cell.ref);
          }
        }
      }
    }

    // 显示所有图片的统计信息
    console.log('\n=== 图片统计信息 ===');
    result.images.forEach((image, imageId) => {
      console.log(`图片 ${imageId}:`);
      console.log(`  - 描述: ${image.description}`);
      console.log(`  - 尺寸: ${image.position.width}x${image.position.height}`);
      console.log(`  - MIME类型: ${image.mimeType}`);
      console.log(`  - Base64数据长度: ${image.base64.length} 字符`);
    });

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

/**
 * 生成HTML预览（数据 + 图片）
 */
async function generateHtmlPreview(result: ExcelParseResult): Promise<void> {
  const outDir = resolveOutputDir();
  const outPath = path.join(outDir, 'preview.html');

  let html = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"/>
<style>
body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial; margin:16px}
.table{border-collapse:collapse; width:100%; margin:12px 0}
.table th,.table td{border:1px solid #e5e7eb; padding:8px; vertical-align:top}
.table th{background:#f8fafc; text-align:left}
.sheet{border:1px solid #e5e7eb; margin-bottom:24px; border-radius:8px; overflow:hidden}
.sheet-header{background:#f1f5f9; padding:12px 16px; font-weight:600}
.image-cell{background:#f8fafc}
.image-cell img{max-width:120px; max-height:120px; display:block}
.meta{color:#475569; font-size:12px}
.badge{display:inline-block; background:#eff6ff; color:#1d4ed8; padding:2px 8px; border-radius:9999px; font-size:12px; margin-left:8px}
</style></head><body>`;

  html += `<h1>Excel 预览</h1>`;
  html += `<div class="meta">工作表数量: ${result.worksheets.length}，图片数量: ${result.images.size}</div>`;

  for (const ws of result.worksheets) {
    html += `<div class="sheet">`;
    html += `<div class="sheet-header">${ws.name}<span class="badge">范围 ${ws.dimension.start} ~ ${ws.dimension.end}</span><span class="badge">图片 ${ws.totalImages}</span></div>`;

    // 头部行（简单使用第一行的列数来渲染表头）
    const maxCols = Math.max(0, ...ws.rows.map(r => r.cells.length));
    html += `<table class="table"><thead><tr>`;
    html += `<th>#</th>`;
    for (let c = 0; c < maxCols; c++) html += `<th>列${c + 1}</th>`;
    html += `</tr></thead><tbody>`;

    // 数据行
    for (const row of ws.rows) {
      html += `<tr>`;
      html += `<td>${row.rowNumber}${row.imageCount > 0 ? ` <span class=\"badge\">📷×${row.imageCount}</span>` : ''}</td>`;

      for (let c = 0; c < maxCols; c++) {
        const cell = row.cells[c];
        if (!cell) { html += `<td></td>`; continue; }

        if (cell.image) {
          html += `<td class="image-cell">`;
          html += `<div><img src="${cell.image.base64}" alt="${cell.image.description}"/></div>`;
          html += `<div class="meta">${cell.image.description || ''}</div>`;
          html += `</td>`;
        } else {
          const text = (cell.type === 'formula' && cell.formula) ? `=${cell.formula}` : String(cell.value || '');
          html += `<td>${escapeHtml(text)}</td>`;
        }
      }

      html += `</tr>`;
    }

    html += `</tbody></table></div>`;
  }

  html += `</body></html>`;
  fs.writeFileSync(outPath, html, 'utf8');
  console.log(`\nHTML预览已生成: ${outPath}`);
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

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
