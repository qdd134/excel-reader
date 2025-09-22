// @ts-ignore
import * as XLSX from 'xlsx';
// @ts-ignore
import * as JSZip from 'jszip';
import { 
  ExcelParseResult, 
  WorksheetData, 
  RowData, 
  CellData, 
  CellImage, 
  ParseOptions, 
  ImageExtractionResult 
} from './types';

/**
 * Excel图片读取器
 * 支持读取Excel文件中的表格数据和嵌入的图片，并将图片转换为base64格式
 */
export class ExcelImageReader {
  private zip: JSZip | null = null;
  private cellImagesXml: string | null = null;
  private cellImagesRels: string | null = null;

  /**
   * 解析Excel文件
   * @param filePath Excel文件路径
   * @param options 解析选项
   * @returns 解析结果
   */
  async parseFile(filePath: string, options: ParseOptions = {}): Promise<ExcelParseResult> {
    const defaultOptions: ParseOptions = {
      includeImages: true,
      imageQuality: 0.8,
      includeEmptyRows: false,
      includeEmptyColumns: false,
      ...options
    };

    const result: ExcelParseResult = {
      worksheets: [],
      images: new Map(),
      errors: []
    };

    try {
      // 读取并解析Excel文件
      const workbook = XLSX.readFile(filePath);
      
      // 如果包含图片，提取图片数据
      if (defaultOptions.includeImages) {
        await this.extractImages(filePath, result);
      }

      // 解析工作表数据
      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const worksheetData = this.parseWorksheet(worksheet, sheetName, result.images, defaultOptions);
        result.worksheets.push(worksheetData);
      }

    } catch (error) {
      result.errors.push(`解析文件失败: ${error instanceof Error ? error.message : String(error)}`);
    }

    return result;
  }

  /**
   * 解析Excel文件Buffer
   * @param buffer Excel文件Buffer
   * @param options 解析选项
   * @returns 解析结果
   */
  async parseBuffer(buffer: any, options: ParseOptions = {}): Promise<ExcelParseResult> {
    const defaultOptions: ParseOptions = {
      includeImages: true,
      imageQuality: 0.8,
      includeEmptyRows: false,
      includeEmptyColumns: false,
      ...options
    };

    const result: ExcelParseResult = {
      worksheets: [],
      images: new Map(),
      errors: []
    };

    try {
      // 读取并解析Excel文件
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      
      // 如果包含图片，提取图片数据
      if (defaultOptions.includeImages) {
        await this.extractImagesFromBuffer(buffer, result);
      }

      // 解析工作表数据
      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const worksheetData = this.parseWorksheet(worksheet, sheetName, result.images, defaultOptions);
        result.worksheets.push(worksheetData);
      }

    } catch (error) {
      result.errors.push(`解析Buffer失败: ${error instanceof Error ? error.message : String(error)}`);
    }

    return result;
  }

  /**
   * 从文件路径提取图片
   */
  private async extractImages(filePath: string, result: ExcelParseResult): Promise<void> {
    try {
      // @ts-ignore
      const fs = await import('fs');
      const fileBuffer = fs.readFileSync(filePath);
      await this.extractImagesFromBuffer(fileBuffer, result);
    } catch (error) {
      result.errors.push(`读取文件失败: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * 从Buffer提取图片
   */
  private async extractImagesFromBuffer(buffer: any, result: ExcelParseResult): Promise<void> {
    try {
      this.zip = await JSZip.loadAsync(buffer, { checkCRC32: false }) as any;
      
      // 读取cellimages.xml文件
      const cellImagesFile = this.zip?.file('xl/cellimages.xml');
      if (cellImagesFile) {
        this.cellImagesXml = await cellImagesFile.async('text');
      }

      // 读取cellimages.xml.rels文件
      const cellImagesRelsFile = this.zip?.file('xl/_rels/cellimages.xml.rels');
      if (cellImagesRelsFile) {
        this.cellImagesRels = await cellImagesRelsFile.async('text');
      }

      // 解析图片信息
      if (this.cellImagesXml && this.cellImagesRels) {
        await this.parseCellImages(result);
      }

    } catch (error) {
      result.errors.push(`提取图片失败: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * 解析cellimages.xml文件
   */
  private async parseCellImages(result: ExcelParseResult): Promise<void> {
    if (!this.cellImagesXml || !this.cellImagesRels || !this.zip) {
      return;
    }

    try {
      // 解析关系映射
      const relationshipMap = this.parseRelationships(this.cellImagesRels);
      
      // 解析图片信息
      const cellImages = this.parseCellImagesXml(this.cellImagesXml);
      
      // 提取每个图片的base64数据
      for (const cellImage of cellImages) {
        const relationship = relationshipMap.get(cellImage.relationshipId);
        if (relationship) {
          const imageResult = await this.extractImageData(relationship.target);
          if (imageResult) {
            const cellImageData: CellImage = {
              id: cellImage.id,
              description: cellImage.description,
              base64: imageResult.base64,
              mimeType: imageResult.mimeType,
              position: cellImage.position,
              relationshipId: cellImage.relationshipId
            };
            result.images.set(cellImage.id, cellImageData);
          }
        }
      }

    } catch (error) {
      result.errors.push(`解析图片信息失败: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * 解析关系文件
   */
  private parseRelationships(relsXml: string): Map<string, { id: string; type: string; target: string }> {
    const relationships = new Map<string, { id: string; type: string; target: string }>();
    
    const relationshipRegex = /<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"\s*\/?>/g;
    let match;
    
    while ((match = relationshipRegex.exec(relsXml)) !== null) {
      relationships.set(match[1], {
        id: match[1],
        type: match[2],
        target: match[3]
      });
    }
    
    return relationships;
  }

  /**
   * 解析cellimages.xml文件
   */
  private parseCellImagesXml(xml: string): Array<{
    id: string;
    description: string;
    position: { x: number; y: number; width: number; height: number };
    relationshipId: string;
  }> {
    const cellImages: Array<{
      id: string;
      description: string;
      position: { x: number; y: number; width: number; height: number };
      relationshipId: string;
    }> = [];

    // 匹配每个cellImage块
    const cellImageRegex = /<etc:cellImage>([\s\S]*?)<\/etc:cellImage>/g;
    let cellImageMatch;

    while ((cellImageMatch = cellImageRegex.exec(xml)) !== null) {
      const cellImageXml = cellImageMatch[1];
      
      // 提取ID和描述
      const idMatch = cellImageXml.match(/name="([^"]+)"/);
      const descMatch = cellImageXml.match(/descr="([^"]+)"/);
      const embedMatch = cellImageXml.match(/r:embed="([^"]+)"/);
      
      // 提取位置信息
      const offMatch = cellImageXml.match(/<a:off\s+x="([^"]+)"\s+y="([^"]+)"\s*\/?>/);
      const extMatch = cellImageXml.match(/<a:ext\s+cx="([^"]+)"\s+cy="([^"]+)"\s*\/?>/);

      if (idMatch && descMatch && embedMatch && offMatch && extMatch) {
        cellImages.push({
          id: idMatch[1],
          description: descMatch[1],
          position: {
            x: parseInt(offMatch[1]),
            y: parseInt(offMatch[2]),
            width: parseInt(extMatch[1]),
            height: parseInt(extMatch[2])
          },
          relationshipId: embedMatch[1]
        });
      }
    }

    return cellImages;
  }

  /**
   * 提取图片数据
   */
  private async extractImageData(imagePath: string): Promise<ImageExtractionResult | null> {
    if (!this.zip) return null;

    try {
      const imageFile = this.zip.file(`xl/${imagePath}`);
      if (!imageFile) return null;

      const imageBuffer = await imageFile.async('nodebuffer');
      const mimeType = this.getMimeTypeFromPath(imagePath);
      const base64 = imageBuffer.toString('base64');

      return {
        id: '',
        description: '',
        base64: `data:${mimeType};base64,${base64}`,
        mimeType,
        rawData: new Uint8Array(imageBuffer)
      };
    } catch (error) {
      return null;
    }
  }

  /**
   * 根据文件路径获取MIME类型
   */
  private getMimeTypeFromPath(path: string): string {
    const extension = path.toLowerCase().split('.').pop();
    switch (extension) {
      case 'jpg':
      case 'jpeg':
        return 'image/jpeg';
      case 'png':
        return 'image/png';
      case 'gif':
        return 'image/gif';
      case 'bmp':
        return 'image/bmp';
      case 'webp':
        return 'image/webp';
      default:
        return 'image/jpeg';
    }
  }

  /**
   * 解析工作表数据
   */
  private parseWorksheet(
    worksheet: XLSX.WorkSheet, 
    sheetName: string, 
    images: Map<string, CellImage>,
    options: ParseOptions
  ): WorksheetData {
    const rows: RowData[] = [];
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    
    // 解析列配置
    const columns = this.parseColumns(worksheet);
    
    // 统计变量
    let totalImages = 0;
    let rowsWithImages = 0;
    
    // 解析行数据
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      const rowData: RowData = {
        rowNumber: rowNum + 1,
        cells: [],
        imageCount: 0,
        imageCells: []
      };

      // 获取行高
      if (worksheet['!rows'] && worksheet['!rows'][rowNum]) {
        const rowInfo = worksheet['!rows'][rowNum];
        if (rowInfo.hpt) {
          rowData.height = rowInfo.hpt;
          rowData.customHeight = true;
        }
      }

      // 解析单元格
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellRef = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellRef];
        
        if (cell || options.includeEmptyColumns) {
          const cellData = this.parseCell(cell, cellRef, images);
          rowData.cells.push(cellData);
          
          // 检查是否为图片单元格
          if (this.isImageCell(cell)) {
            const imageId = this.extractImageIdFromCell(cell);
            if (imageId && images.has(imageId)) {
              rowData.imageCount++;
              rowData.imageCells.push(cellRef);
              totalImages++;
            }
          }
        }
      }

      // 如果该行包含图片，增加计数
      if (rowData.imageCount > 0) {
        rowsWithImages++;
      }

      // 如果包含空行或行中有数据，添加到结果中
      if (options.includeEmptyRows || rowData.cells.some(cell => cell.value !== '')) {
        rows.push(rowData);
      }
    }

    return {
      name: sheetName,
      dimension: {
        start: XLSX.utils.encode_cell(range.s),
        end: XLSX.utils.encode_cell(range.e)
      },
      rows,
      columns,
      totalImages,
      rowsWithImages
    };
  }

  /**
   * 解析列配置
   */
  private parseColumns(worksheet: XLSX.WorkSheet): Array<{
    min: number;
    max: number;
    width: number;
    customWidth: boolean;
  }> {
    const columns: Array<{
      min: number;
      max: number;
      width: number;
      customWidth: boolean;
    }> = [];

    if (worksheet['!cols']) {
      for (const col of worksheet['!cols']) {
        columns.push({
          min: (col as any).min || 0,
          max: (col as any).max || 0,
          width: col.width || 9,
          customWidth: (col as any).customWidth || false
        });
      }
    }

    return columns;
  }

  /**
   * 解析单元格数据
   */
  private parseCell(cell: XLSX.CellObject | undefined, cellRef: string, images: Map<string, CellImage>): CellData {
    const cellData: CellData = {
      ref: cellRef,
      value: '',
      type: 'string'
    };

    if (!cell) {
      return cellData;
    }

    // 设置基本属性
    cellData.value = String(cell.v || '');
    cellData.styleId = cell.s;
    
    // 判断单元格类型
    if (cell.t === 's') {
      cellData.type = 'string';
    } else if (cell.t === 'n') {
      cellData.type = 'number';
    } else if (cell.f) {
      cellData.type = 'formula';
      cellData.formula = cell.f;
    }
    
    // 检查是否包含DISPIMG公式（无论单元格类型如何）
    if (typeof cell.v === 'string' && cell.v.includes('DISPIMG')) {
      cellData.type = 'image';
      // 尝试从公式中提取图片ID
      const imageIdMatch = cell.v.match(/DISPIMG\("([^"]+)"/);
      if (imageIdMatch) {
        const imageId = imageIdMatch[1];
        const image = images.get(imageId);
        if (image) {
          cellData.image = image;
        }
      }
    } else if (cell.f && cell.f.includes('DISPIMG')) {
      cellData.type = 'image';
      // 尝试从公式中提取图片ID
      const imageIdMatch = cell.f.match(/DISPIMG\("([^"]+)"/);
      if (imageIdMatch) {
        const imageId = imageIdMatch[1];
        const image = images.get(imageId);
        if (image) {
          cellData.image = image;
        }
      }
    }

    return cellData;
  }

  /**
   * 检查单元格是否包含图片
   */
  private isImageCell(cell: XLSX.CellObject | undefined): boolean {
    if (!cell) return false;
    
    // 检查单元格值是否包含DISPIMG公式
    if (typeof cell.v === 'string' && cell.v.includes('DISPIMG')) {
      return true;
    }
    
    // 检查公式是否包含DISPIMG
    if (cell.f && cell.f.includes('DISPIMG')) {
      return true;
    }
    
    return false;
  }

  /**
   * 从单元格中提取图片ID
   */
  private extractImageIdFromCell(cell: XLSX.CellObject | undefined): string | null {
    if (!cell) return null;
    
    let textToSearch = '';
    
    // 优先从公式中提取
    if (cell.f && cell.f.includes('DISPIMG')) {
      textToSearch = cell.f;
    } else if (typeof cell.v === 'string' && cell.v.includes('DISPIMG')) {
      textToSearch = cell.v;
    }
    
    if (textToSearch) {
      const imageIdMatch = textToSearch.match(/DISPIMG\("([^"]+)"/);
      return imageIdMatch ? imageIdMatch[1] : null;
    }
    
    return null;
  }
}
