require('dotenv').config({ path: require('path').resolve(__dirname, '.env') });
const { ExcelImageReader } = require('../dist/index');
const fs = require('fs');
const path = require('path');

/**
 * 简化的示例：读取Excel文件并提取图片数据
 */
async function main() {
  const reader = new ExcelImageReader();
  
  // 示例文件路径（请替换为实际的Excel文件路径）
  const envPath = process.env.EXCEL_READER_XLSX || process.env.EXCEL_XLSX || process.env.XLSX_PATH;
  const defaultPath = path.join(__dirname, 'test.xlsx');
  const excelFilePath = (envPath && fs.existsSync(envPath)) ? envPath : defaultPath;
  
  try {
    console.log('开始解析Excel文件...');
    
    // 解析Excel文件
    const result = await reader.parseFile(excelFilePath, {
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
async function saveImageToFile(image, worksheetName, cellRef) {
  try {
    const outputBase = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    const outputDir = outputBase;
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

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
 * 根据MIME类型获取文件扩展名
 */
function getFileExtension(mimeType) {
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

// 运行示例
if (require.main === module) {
  main().then(() => {
    console.log('\n示例执行完成！');
  }).catch(error => {
    console.error('示例执行失败:', error);
  });
}

module.exports = { main };
