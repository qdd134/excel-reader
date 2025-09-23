import * as JSZip from 'jszip';
import { CellImage, ExcelParseResult } from './types';
import { DrawingParser, FloatingImageInfo } from './DrawingParser';
import { RelationshipParser, RelationshipInfo } from './RelationshipParser';
import { ImageExtractor } from './ImageExtractor';

/**
 * 浮动图片管理器
 * 负责处理浮动图片的解析、绑定和归类
 */
export class FloatingImageManager {
  private drawingParser: DrawingParser;
  private relationshipParser: RelationshipParser;
  private imageExtractor: ImageExtractor;
  private floatingImagesBySheet: Map<string, Map<string, string[]>> = new Map();

  constructor() {
    this.drawingParser = new DrawingParser();
    this.relationshipParser = new RelationshipParser();
    this.imageExtractor = new ImageExtractor();
  }

  /**
   * 设置JSZip实例
   */
  setZip(zip: JSZip): void {
    this.imageExtractor.setZip(zip);
  }

  /**
   * 解析所有浮动图片
   */
  async parseFloatingImages(result: ExcelParseResult): Promise<void> {
    try {
      // 获取工作表映射
      const sheetsMap = await this.getSheetsMap();
      
      // 遍历每个工作表
      for (const [sheetName, sheetPath] of sheetsMap) {
        await this.parseSheetFloatingImages(sheetName, sheetPath, result);
      }
    } catch (error) {
      console.error('解析浮动图片失败:', error);
      result.errors.push(`解析浮动图片失败: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * 获取工作表名称到路径的映射
   */
  private async getSheetsMap(): Promise<Map<string, string>> {
    const zip = (this.imageExtractor as any).zip as JSZip;
    if (!zip) return new Map();

    const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
    const workbookRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
    
    if (!workbookXml || !workbookRelsXml) return new Map();
    
    return this.relationshipParser.mapSheetNamesToPaths(workbookXml, workbookRelsXml);
  }

  /**
   * 解析单个工作表的浮动图片
   */
  private async parseSheetFloatingImages(
    sheetName: string, 
    sheetPath: string, 
    result: ExcelParseResult
  ): Promise<void> {
    const zip = (this.imageExtractor as any).zip as JSZip;
    if (!zip) return;

    try {
      // 获取工作表关系文件
      const sheetRels = await this.getSheetRelationships(sheetPath, zip);
      if (!sheetRels) return;

      // 查找drawing关系
      const drawingRel = this.relationshipParser.findDrawingRelationship(sheetRels);
      if (!drawingRel) return;

      // 解析drawing文件
      const floatingImages = await this.parseDrawingFile(drawingRel, zip);
      
      // 处理每个浮动图片
      for (const imageInfo of floatingImages) {
        await this.processFloatingImage(imageInfo, sheetName, result);
      }
    } catch (error) {
      console.error(`解析工作表 ${sheetName} 的浮动图片失败:`, error);
    }
  }

  /**
   * 获取工作表关系文件
   */
  private async getSheetRelationships(sheetPath: string, zip: JSZip): Promise<Map<string, RelationshipInfo> | null> {
    const baseName = sheetPath.split('/').pop() as string;
    const sheetRelsPath = `xl/worksheets/_rels/${baseName}.rels`;
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async('text');
    
    if (!sheetRelsXml) return null;
    
    return this.relationshipParser.parseRelationships(sheetRelsXml);
  }

  /**
   * 解析drawing文件
   */
  private async parseDrawingFile(drawingRel: RelationshipInfo, zip: JSZip): Promise<FloatingImageInfo[]> {
    // 读取drawing xml
    const drawingTarget = this.relationshipParser.normalizeTargetPath('worksheets', drawingRel.target);
    const drawingXml = await zip.file(`xl/${drawingTarget}`)?.async('text');
    if (!drawingXml) return [];

    // 读取drawing关系文件
    const drawingRelsBase = drawingTarget.replace('drawings/', 'drawings/_rels/');
    const drawingRelsPath = `xl/${drawingRelsBase}.rels`;
    const drawingRelsXml = await zip.file(drawingRelsPath)?.async('text');
    const drawingRels = drawingRelsXml ? this.relationshipParser.parseRelationships(drawingRelsXml) : new Map();

    // 解析drawing xml
    const floatingImages = this.drawingParser.parseDrawingXml(drawingXml);
    
    // 为每个浮动图片添加媒体路径信息
    for (const imageInfo of floatingImages) {
      const rel = drawingRels.get(imageInfo.relationshipId);
      if (!rel) {
        (imageInfo as any).__skip = 'missing_relationship';
        continue;
      }
      // skip External links
      if (rel.targetMode && String(rel.targetMode).toLowerCase() === 'external') {
        (imageInfo as any).__skip = 'external';
        continue;
      }
      const mediaTarget = this.relationshipParser.normalizeTargetPath('drawings', rel.target);
      (imageInfo as any).mediaPath = mediaTarget;
    }

    return floatingImages;
  }

  /**
   * 处理单个浮动图片
   */
  private async processFloatingImage(
    imageInfo: FloatingImageInfo & { mediaPath?: string },
    sheetName: string,
    result: ExcelParseResult
  ): Promise<void> {
    if ((imageInfo as any).__skip) {
      (result.errors || []).push(`[skip] reason=${(imageInfo as any).__skip} id=${imageInfo.id} sheet=${sheetName} cell=${imageInfo.cellRef}`);
      return;
    }
    if (!imageInfo.mediaPath) {
      (result.errors || []).push(`[skip] reason=missing_media_path id=${imageInfo.id} sheet=${sheetName} cell=${imageInfo.cellRef}`);
      return;
    }

    try {
      // 提取图片数据
      const imageResult = await this.imageExtractor.extractImageData(imageInfo.mediaPath);
      if (!imageResult) {
        (result.errors || []).push(`[skip] reason=extract_failed id=${imageInfo.id} sheet=${sheetName} cell=${imageInfo.cellRef}`);
        return;
      }

      // 创建CellImage对象
      const cellImage = this.imageExtractor.createCellImage(
        imageInfo.id,
        imageInfo.description,
        imageInfo.relationshipId,
        imageInfo.position,
        imageResult
      );

      // 添加到结果中
      if (!result.images.has(imageInfo.id)) {
        result.images.set(imageInfo.id, cellImage);
      }

      // 绑定到单元格
      this.bindImageToCell(imageInfo, sheetName);
    } catch (error) {
      console.error(`处理浮动图片 ${imageInfo.id} 失败:`, error);
    }
  }

  /**
   * 将图片绑定到单元格
   */
  private bindImageToCell(imageInfo: FloatingImageInfo, sheetName: string): void {
    if (!this.floatingImagesBySheet.has(sheetName)) {
      this.floatingImagesBySheet.set(sheetName, new Map());
    }
    
    const byCell = this.floatingImagesBySheet.get(sheetName)!;
    if (!byCell.has(imageInfo.cellRef)) {
      byCell.set(imageInfo.cellRef, []);
    }
    
    byCell.get(imageInfo.cellRef)!.push(imageInfo.id);
  }

  /**
   * 获取指定工作表和单元格的浮动图片ID列表
   */
  getFloatingImageIds(sheetName: string, cellRef: string): string[] {
    const sheetImages = this.floatingImagesBySheet.get(sheetName);
    if (!sheetImages) return [];
    
    return sheetImages.get(cellRef) || [];
  }

  /**
   * 获取指定工作表的所有浮动图片绑定信息
   */
  getSheetFloatingImages(sheetName: string): Map<string, string[]> {
    return this.floatingImagesBySheet.get(sheetName) || new Map();
  }

  /**
   * 获取所有浮动图片绑定信息
   */
  getAllFloatingImages(): Map<string, Map<string, string[]>> {
    return this.floatingImagesBySheet;
  }

  /**
   * 清空浮动图片数据
   */
  clear(): void {
    this.floatingImagesBySheet.clear();
  }
}
