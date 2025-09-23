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
import { XMLParser } from 'fast-xml-parser';
import { FloatingImageManager } from './FloatingImageManager';
import { RelationshipParser } from './RelationshipParser';
import { ImageExtractor } from './ImageExtractor';

/**
 * Excel图片读取器
 * 支持读取Excel文件中的表格数据和嵌入的图片，并将图片转换为base64格式
 */
export class ExcelImageReader {
  private zip: JSZip | null = null;
  private cellImagesXml: string | null = null;
  private cellImagesRels: string | null = null;
  private floatingImageManager: FloatingImageManager;
  private relationshipParser: RelationshipParser;
  private imageExtractor: ImageExtractor;

  constructor() {
    this.floatingImageManager = new FloatingImageManager();
    this.relationshipParser = new RelationshipParser();
    this.imageExtractor = new ImageExtractor();
  }

  /**
   * 解析Excel文件
   * @param filePath Excel文件路径
   * @param options 解析选项
   * @returns 解析结果
   */
  async parseFile(filePath: string, options: ParseOptions = {}): Promise<ExcelParseResult> {
    // 重置跨文件状态
    this.zip = null;
    this.cellImagesXml = null;
    this.cellImagesRels = null;
    if (this.floatingImageManager) this.floatingImageManager.clear();
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
    // 重置跨文件状态
    this.zip = null;
    this.cellImagesXml = null;
    this.cellImagesRels = null;
    if (this.floatingImageManager) this.floatingImageManager.clear();
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
      
      // 设置组件依赖
      if (this.zip) {
        this.floatingImageManager.setZip(this.zip);
        this.imageExtractor.setZip(this.zip);
      }
      
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

      // 解析浮动图片
      await this.floatingImageManager.parseFloatingImages(result);
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
      const relationshipMap = this.relationshipParser.parseRelationships(this.cellImagesRels);
      
      // 解析图片信息
      const cellImages = this.parseCellImagesXml(this.cellImagesXml);
      
      // console.log("relationshipMap:",relationshipMap,"cellImages:", cellImages);
      // 提取每个图片的base64数据
      for (const cellImage of cellImages) {
        const relationship = relationshipMap.get(cellImage.relationshipId);
        if (relationship) {
          const imageResult = await this.imageExtractor.extractImageData(relationship.target);
          if (imageResult) {
            const cellImageData = this.imageExtractor.createCellImage(
              cellImage.id,
              cellImage.description,
              cellImage.relationshipId,
              cellImage.position,
              imageResult
            );
            result.images.set(cellImage.id, cellImageData);
          }
        }
      }

    } catch (error) {
      result.errors.push(`解析图片信息失败: ${error instanceof Error ? error.message : String(error)}`);
    }
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
    const cellImages: Array<{ id: string; description: string; position: { x: number; y: number; width: number; height: number }; relationshipId: string; }> = [];
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
    try {
      const doc = parser.parse(xml);
      const root = (doc['etc:cellImages'] || doc['cellImages'] || doc);
      const itemsRaw = root?.['etc:cellImage'] || root?.['cellImage'] || [];
      const items = Array.isArray(itemsRaw) ? itemsRaw : [itemsRaw];
      for (const item of items) {
        if (!item) continue;
        const pic = item['xdr:pic'] || item['pic'] || item.pic;
        if (!pic) continue;
        const nvPicPr = pic['xdr:nvPicPr'] || pic['nvPicPr'];
        const cNvPr = nvPicPr?.['xdr:cNvPr'] || nvPicPr?.['cNvPr'] || nvPicPr?.cNvPr;
        const blipFill = pic['xdr:blipFill'] || pic['blipFill'];
        const aBlip = blipFill?.['a:blip'] || blipFill?.blip;
        const spPr = pic['xdr:spPr'] || pic['spPr'];
        const xfrm = spPr?.['a:xfrm'] || spPr?.xfrm;
        const off = xfrm?.['a:off'] || xfrm?.off;
        const ext = xfrm?.['a:ext'] || xfrm?.ext;

        const id = cNvPr?.name;
        const description = cNvPr?.descr ?? '';
        const relationshipId = aBlip?.['r:embed'] || aBlip?.embed;
        const x = Number(off?.x ?? 0);
        const y = Number(off?.y ?? 0);
        const width = Number(ext?.cx ?? 0);
        const height = Number(ext?.cy ?? 0);

        if (id && relationshipId && Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(width) && Number.isFinite(height)) {
          cellImages.push({
            id,
            description,
            position: { x, y, width, height },
            relationshipId
          });
        }
      }
    } catch {
      // ignore
    }
    return cellImages;
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
    const floatingForSheet = this.floatingImageManager.getSheetFloatingImages(sheetName);

    // 扩展数据范围以包含浮动图片所在的单元格（避免因插入表头导致图片行超出 !ref 而被忽略）
    let startRow = range.s.r;
    let startCol = range.s.c;
    let endRow = range.e.r;
    let endCol = range.e.c;
    for (const cellRef of Array.from(floatingForSheet.keys())) {
      const coord = XLSX.utils.decode_cell(cellRef);
      if (Number.isFinite(coord.r) && Number.isFinite(coord.c)) {
        if (coord.r > endRow) endRow = coord.r;
        if (coord.c > endCol) endCol = coord.c;
        if (coord.r < startRow) startRow = coord.r;
        if (coord.c < startCol) startCol = coord.c;
      }
    }
    
    // 解析行数据
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const rowSeen = new Set<string>();
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
      for (let colNum = startCol; colNum <= endCol; colNum++) {
        const cellRef = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellRef];

        // 先处理浮动图计数（即使该列没有实际单元格也要统计）
        const floatingIds = floatingForSheet.get(cellRef);
        if (floatingIds && floatingIds.length) {
          // de-dup per row / global by image id
          for (const fid of floatingIds) {
            if (!rowSeen.has(fid)) {
              rowSeen.add(fid);
              rowData.imageCount += 1;
              totalImages += 1;
              rowData.imageCells.push(cellRef);
            }
          }
        }

        if (cell || options.includeEmptyColumns) {
          const cellData = this.parseCell(cell, cellRef, images);
          // 若该单元格还没有 image 字段，则选择第一个作为该单元格代表图片
          if (!cellData.image && floatingIds && floatingIds.length) {
            const img = images.get(floatingIds[0]);
            if (img) cellData.image = img;
          }
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

    // 设置基本属性（保留数字 0，不被当成空）
    cellData.value = String(cell.v ?? '');
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
