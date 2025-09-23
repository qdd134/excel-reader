// @ts-nocheck
import { ExcelImageReader, ExcelParseResult, CellData, WorksheetData } from '../src/index';
import dotenv from 'dotenv';
dotenv.config({ path: require('path').resolve(__dirname, '.env') });
import * as fs from 'fs';
import * as path from 'path';

/**
 * 高级示例：处理复杂的Excel文件
 */
class AdvancedExcelProcessor {
  private reader: ExcelImageReader;

  constructor() {
    this.reader = new ExcelImageReader();
  }

  /**
   * 处理Excel文件并生成JSON报告
   */
  async processExcelFile(filePath: string): Promise<void> {
    console.log(`开始处理文件: ${filePath}`);
    
    const result = await this.reader.parseFile(filePath, {
      includeImages: true,
      includeEmptyRows: false,
      includeEmptyColumns: false
    });

    // 生成详细报告
    const report = this.generateReport(result);
    
    // 保存报告到文件
    const outputBase = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    const reportPath = path.join(outputBase, 'excel_report.json');
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(`报告已保存到: ${reportPath}`);

    // 提取所有图片
    await this.extractAllImages(result);

    // 生成CSV文件
    await this.generateCSV(result);

    // 生成HTML预览
    await this.generateHTMLPreview(result);
  }

  /**
   * 生成详细报告
   */
  private generateReport(result: ExcelParseResult): any {
    const report = {
      summary: {
        totalWorksheets: result.worksheets.length,
        totalImages: result.images.size,
        totalRows: 0,
        totalCells: 0,
        errors: result.errors.length
      },
      worksheets: [] as any[],
      images: [] as any[],
      errors: result.errors
    };

    // 统计工作表信息
    for (const worksheet of result.worksheets) {
      const worksheetInfo = {
        name: worksheet.name,
        dimension: worksheet.dimension,
        rowCount: worksheet.rows.length,
        columnCount: worksheet.columns.length,
        imageCount: 0,
        dataTypes: {
          string: 0,
          number: 0,
          formula: 0,
          image: 0
        }
      };

      // 统计单元格类型
      for (const row of worksheet.rows) {
        for (const cell of row.cells) {
          worksheetInfo.dataTypes[cell.type]++;
          if (cell.image) {
            worksheetInfo.imageCount++;
          }
        }
      }

      report.worksheets.push(worksheetInfo);
      report.summary.totalRows += worksheet.rows.length;
      report.summary.totalCells += worksheetInfo.dataTypes.string + 
                                   worksheetInfo.dataTypes.number + 
                                   worksheetInfo.dataTypes.formula + 
                                   worksheetInfo.dataTypes.image;
    }

    // 统计图片信息
    result.images.forEach((image, imageId) => {
      report.images.push({
        id: imageId,
        description: image.description,
        mimeType: image.mimeType,
        dimensions: {
          width: image.position.width,
          height: image.position.height
        },
        position: image.position,
        base64Length: image.base64.length
      });
    });

    return report;
  }

  /**
   * 提取所有图片
   */
  private async extractAllImages(result: ExcelParseResult): Promise<void> {
    const outputBase = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    const outputDir = path.join(outputBase, 'images');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log(`开始提取 ${result.images.size} 张图片...`);

    for (const [imageId, image] of result.images) {
      const fileName = `${imageId}.${this.getFileExtension(image.mimeType)}`;
      const filePath = path.join(outputDir, fileName);
      
      // 提取base64数据
      const base64Data = image.base64.split(',')[1];
      const buffer = Buffer.from(base64Data, 'base64');
      
      fs.writeFileSync(filePath, buffer);
      console.log(`图片已保存: ${fileName}`);
    }

    console.log(`所有图片已保存到: ${outputDir}`);
  }

  /**
   * 生成CSV文件
   */
  private async generateCSV(result: ExcelParseResult): Promise<void> {
    const outputDir = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    for (const worksheet of result.worksheets) {
      const csvPath = path.join(outputDir, `${worksheet.name}.csv`);
      const csvLines: string[] = [];

      for (const row of worksheet.rows) {
        const csvRow: string[] = [];
        
        for (const cell of row.cells) {
          let cellValue = '';
          
          if (cell.type === 'image' && cell.image) {
            cellValue = `[图片: ${cell.image.description}]`;
          } else if (cell.type === 'formula' && cell.formula) {
            cellValue = `=${cell.formula}`;
          } else {
            cellValue = String(cell.value || '');
          }
          
          // 转义CSV特殊字符
          if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
            cellValue = `"${cellValue.replace(/"/g, '""')}"`;
          }
          
          csvRow.push(cellValue);
        }
        
        csvLines.push(csvRow.join(','));
      }

      fs.writeFileSync(csvPath, csvLines.join('\n'), 'utf8');
      console.log(`CSV文件已保存: ${csvPath}`);
    }
  }

  /**
   * 生成HTML预览
   */
  private async generateHTMLPreview(result: ExcelParseResult): Promise<void> {
    const outputDir = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const htmlPath = path.join(outputDir, 'preview.html');
    let html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel文件预览</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .worksheet { margin-bottom: 40px; border: 1px solid #ddd; padding: 20px; }
        .worksheet h2 { color: #333; border-bottom: 2px solid #007acc; padding-bottom: 10px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .image-cell { background-color: #e8f4fd; }
        .image-cell img { max-width: 100px; max-height: 100px; }
        .summary { background-color: #f9f9f9; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>Excel文件预览</h1>
    <div class="summary">
        <h3>文件摘要</h3>
        <p>工作表数量: ${result.worksheets.length}</p>
        <p>图片数量: ${result.images.size}</p>
        <p>总行数: ${result.worksheets.reduce((sum, ws) => sum + ws.rows.length, 0)}</p>
    </div>
`;

    // 为每个工作表生成HTML表格
    for (const worksheet of result.worksheets) {
      html += `
    <div class="worksheet">
        <h2>${worksheet.name}</h2>
        <p>数据范围: ${worksheet.dimension.start}:${worksheet.dimension.end}</p>
        <table>
            <thead>
                <tr>
                    <th>行号</th>
                    <th>列A</th>
                    <th>列B</th>
                    <th>列C</th>
                    <th>列D</th>
                    <th>列E</th>
                </tr>
            </thead>
            <tbody>
`;

      for (const row of worksheet.rows) {
        html += `                <tr>\n`;
        html += `                    <td>${row.rowNumber}</td>\n`;
        
        // 为每列生成单元格
        for (let col = 0; col < 5; col++) {
          const cell = row.cells[col];
          if (cell) {
            if (cell.type === 'image' && cell.image) {
              html += `                    <td class="image-cell">\n`;
              html += `                        <img src="${cell.image.base64}" alt="${cell.image.description}" />\n`;
              html += `                        <br><small>${cell.image.description}</small>\n`;
              html += `                    </td>\n`;
            } else {
              html += `                    <td>${cell.value || ''}</td>\n`;
            }
          } else {
            html += `                    <td></td>\n`;
          }
        }
        
        html += `                </tr>\n`;
      }

      html += `            </tbody>\n        </table>\n    </div>\n`;
    }

    html += `</body>\n</html>`;

    fs.writeFileSync(htmlPath, html, 'utf8');
    console.log(`HTML预览已保存: ${htmlPath}`);
  }

  /**
   * 获取文件扩展名
   */
  private getFileExtension(mimeType: string): string {
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
}

// 运行高级示例
async function runAdvancedExample() {
  const processor = new AdvancedExcelProcessor();
  
  // 优先从环境变量读取路径
  const envPath = process.env.EXCEL_READER_XLSX || process.env.EXCEL_XLSX || process.env.XLSX_PATH;
  const defaultPath = path.join(__dirname, '../test.xlsx');
  const excelFilePath = (envPath && fs.existsSync(envPath)) ? envPath : defaultPath;
  
  if (fs.existsSync(excelFilePath)) {
    await processor.processExcelFile(excelFilePath);
    console.log('高级示例执行完成！');
  } else {
    console.log('请将test.xlsx文件放在example目录下');
  }
}

if (require.main === module) {
  runAdvancedExample().catch(console.error);
}

export { AdvancedExcelProcessor, runAdvancedExample };
