// @ts-nocheck
import { ExcelImageReader, ExcelParseResult, RowData, CellData } from '../src/index';
import dotenv from 'dotenv';
dotenv.config({ path: require('path').resolve(__dirname, '.env') });
import * as fs from 'fs';
import * as path from 'path';

/**
 * 多图片处理示例
 * 演示如何处理一行中包含多个图片的Excel文件
 */
class MultiImageProcessor {
  private reader: ExcelImageReader;

  constructor() {
    this.reader = new ExcelImageReader();
  }

  /**
   * 处理包含多图片的Excel文件
   */
  async processMultiImageExcel(filePath: string): Promise<void> {
    console.log('开始处理多图片Excel文件...');
    
    const result = await this.reader.parseFile(filePath, {
      includeImages: true,
      includeEmptyRows: false,
      includeEmptyColumns: false
    });

    console.log(`解析完成！发现 ${result.worksheets.length} 个工作表`);
    console.log(`发现 ${result.images.size} 张图片`);

    // 分析每个工作表的图片分布
    for (const worksheet of result.worksheets) {
      await this.analyzeWorksheetImages(worksheet);
    }

    // 生成多图片报告
    await this.generateMultiImageReport(result);
  }

  /**
   * 分析工作表的图片分布
   */
  private async analyzeWorksheetImages(worksheet: any): Promise<void> {
    console.log(`\n=== 分析工作表: ${worksheet.name} ===`);
    console.log(`总图片数: ${worksheet.totalImages}`);
    console.log(`包含图片的行数: ${worksheet.rowsWithImages}`);

    // 统计每行的图片数量分布
    const imageDistribution = new Map<number, number>();
    const multiImageRows: RowData[] = [];

    for (const row of worksheet.rows) {
      if (row.imageCount > 0) {
        imageDistribution.set(row.imageCount, (imageDistribution.get(row.imageCount) || 0) + 1);
        
        if (row.imageCount > 1) {
          multiImageRows.push(row);
        }
      }
    }

    // 显示图片分布统计
    console.log('\n📊 图片分布统计:');
    for (const [imageCount, rowCount] of imageDistribution) {
      console.log(`  ${imageCount}张图片的行: ${rowCount}行`);
    }

    // 详细分析多图片行
    if (multiImageRows.length > 0) {
      console.log(`\n🔍 发现 ${multiImageRows.length} 行包含多张图片:`);
      
      for (const row of multiImageRows) {
        console.log(`\n--- 第 ${row.rowNumber} 行 (${row.imageCount}张图片) ---`);
        console.log(`图片位置: ${row.imageCells.join(', ')}`);
        
        // 显示每张图片的详细信息
        for (const cell of row.cells) {
          if (cell.image) {
            console.log(`  📷 ${cell.ref}: ${cell.image.description}`);
            console.log(`     尺寸: ${cell.image.position.width}x${cell.image.position.height}`);
            console.log(`     位置: (${cell.image.position.x}, ${cell.image.position.y})`);
          }
        }
      }
    }

    // 分析图片在列中的分布
    await this.analyzeColumnImageDistribution(worksheet);
  }

  /**
   * 分析图片在列中的分布
   */
  private async analyzeColumnImageDistribution(worksheet: any): Promise<void> {
    console.log('\n📈 列图片分布分析:');
    
    const columnImageCount = new Map<string, number>();
    const columnImageTypes = new Map<string, Set<string>>();

    for (const row of worksheet.rows) {
      for (const cell of row.cells) {
        if (cell.image) {
          const col = cell.ref.replace(/\d+/, ''); // 提取列字母
          columnImageCount.set(col, (columnImageCount.get(col) || 0) + 1);
          
          if (!columnImageTypes.has(col)) {
            columnImageTypes.set(col, new Set());
          }
          columnImageTypes.get(col)!.add(cell.image.mimeType);
        }
      }
    }

    // 显示列图片统计
    for (const [col, count] of columnImageCount) {
      const types = Array.from(columnImageTypes.get(col) || []);
      console.log(`  列 ${col}: ${count}张图片 (类型: ${types.join(', ')})`);
    }
  }

  /**
   * 生成多图片报告
   */
  private async generateMultiImageReport(result: ExcelParseResult): Promise<void> {
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const report = {
      summary: {
        totalWorksheets: result.worksheets.length,
        totalImages: result.images.size,
        totalRows: result.worksheets.reduce((sum, ws) => sum + ws.rows.length, 0),
        rowsWithImages: result.worksheets.reduce((sum, ws) => sum + ws.rowsWithImages, 0),
        multiImageRows: 0
      },
      worksheets: [] as any[],
      imageAnalysis: {
        byRow: new Map<number, number>(),
        byColumn: new Map<string, number>(),
        byType: new Map<string, number>()
      }
    };

    // 分析每个工作表
    for (const worksheet of result.worksheets) {
      const worksheetReport = {
        name: worksheet.name,
        totalImages: worksheet.totalImages,
        rowsWithImages: worksheet.rowsWithImages,
        multiImageRows: 0,
        imageDistribution: new Map<number, number>(),
        columnDistribution: new Map<string, number>()
      };

      // 分析行图片分布
      for (const row of worksheet.rows) {
        if (row.imageCount > 0) {
          worksheetReport.imageDistribution.set(
            row.imageCount, 
            (worksheetReport.imageDistribution.get(row.imageCount) || 0) + 1
          );
          
          if (row.imageCount > 1) {
            worksheetReport.multiImageRows++;
            report.summary.multiImageRows++;
          }
        }

        // 分析列图片分布
        for (const cell of row.cells) {
          if (cell.image) {
            const col = cell.ref.replace(/\d+/, '');
            worksheetReport.columnDistribution.set(
              col, 
              (worksheetReport.columnDistribution.get(col) || 0) + 1
            );
          }
        }
      }

      report.worksheets.push(worksheetReport);
    }

    // 分析图片类型分布
    for (const [imageId, image] of result.images) {
      const mimeType = image.mimeType;
      report.imageAnalysis.byType.set(
        mimeType, 
        (report.imageAnalysis.byType.get(mimeType) || 0) + 1
      );
    }

    // 保存报告
    const reportPath = path.join(outputDir, 'multi_image_report.json');
    fs.writeFileSync(reportPath, JSON.stringify(report, (key, value) => {
      if (value instanceof Map) {
        return Object.fromEntries(value);
      }
      return value;
    }, 2));

    console.log(`\n📋 多图片报告已保存到: ${reportPath}`);
  }

  /**
   * 提取所有多图片行的图片
   */
  async extractMultiImageRows(filePath: string): Promise<void> {
    const result = await this.reader.parseFile(filePath, { includeImages: true });
    const outputBase = process.env.EXCEL_READER_OUTPUT || process.env.EXCEL_OUTPUT || process.env.OUTPUT_DIR || path.join(__dirname, 'output');
    const outputDir = path.join(outputBase, 'multi_image_rows');
    
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    for (const worksheet of result.worksheets) {
      const multiImageRows = worksheet.rows.filter(row => row.imageCount > 1);
      
      if (multiImageRows.length > 0) {
        console.log(`\n📁 提取工作表 ${worksheet.name} 中的多图片行...`);
        
        for (const row of multiImageRows) {
          const rowDir = path.join(outputDir, `${worksheet.name}_row_${row.rowNumber}`);
          fs.mkdirSync(rowDir, { recursive: true });
          
          console.log(`  第 ${row.rowNumber} 行: ${row.imageCount}张图片`);
          
          for (const cell of row.cells) {
            if (cell.image) {
              const fileName = `${cell.ref}_${cell.image.id}.${this.getFileExtension(cell.image.mimeType)}`;
              const filePath = path.join(rowDir, fileName);
              
              const base64Data = cell.image.base64.split(',')[1];
              const buffer = Buffer.from(base64Data, 'base64');
              fs.writeFileSync(filePath, buffer);
              
              console.log(`    ✅ 保存: ${fileName}`);
            }
          }
        }
      }
    }
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

// 运行多图片处理示例
async function runMultiImageExample() {
  const processor = new MultiImageProcessor();
  
  // 优先从环境变量读取路径
  const envPath = process.env.EXCEL_READER_XLSX || process.env.EXCEL_XLSX || process.env.XLSX_PATH;
  const defaultPath = path.join(__dirname, '../test.xlsx');
  const excelFilePath = (envPath && fs.existsSync(envPath)) ? envPath : defaultPath;
  
  if (fs.existsSync(excelFilePath)) {
    await processor.processMultiImageExcel(excelFilePath);
    await processor.extractMultiImageRows(excelFilePath);
    console.log('\n✅ 多图片处理示例执行完成！');
  } else {
    console.log(`文件不存在: ${excelFilePath}`);
    console.log('请将test.xlsx文件放在example目录下');
  }
}

if (require.main === module) {
  runMultiImageExample().catch(console.error);
}

export { MultiImageProcessor, runMultiImageExample };
