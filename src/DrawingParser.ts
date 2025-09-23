import { XMLParser } from 'fast-xml-parser';
import * as XLSX from 'xlsx';

/**
 * Drawing.xml解析器
 * 负责解析Excel中的drawing.xml文件，提取浮动图片信息
 */
export class DrawingParser {
  private parser: XMLParser;

  constructor() {
    this.parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  }

  /**
   * 解析drawing.xml文件，提取浮动图片信息
   */
  parseDrawingXml(drawingXml: string): FloatingImageInfo[] {
    const floatingImages: FloatingImageInfo[] = [];
    
    try {
      const doc = this.parser.parse(drawingXml);
      const wsDr = doc['xdr:wsDr'] || doc['wsDr'] || doc;
      
      // 合并twoCellAnchor和oneCellAnchor
      let anchorsRaw = ([] as any[])
        .concat(wsDr?.['xdr:twoCellAnchor'] || wsDr?.twoCellAnchor || [])
        .concat(wsDr?.['xdr:oneCellAnchor'] || wsDr?.oneCellAnchor || []);
      // unwrap mc:AlternateContent if present
      anchorsRaw = anchorsRaw.map((anc: any) => {
        if (!anc) return anc;
        const ac = anc['mc:AlternateContent'];
        if (!ac) return anc;
        const choice = ac['mc:Choice'] || ac.mcChoice;
        const fallback = ac['mc:Fallback'] || ac.mcFallback;
        // prefer choice, fallback to fallback
        const pic = choice?.['xdr:pic'] || fallback?.['xdr:pic'] || choice?.pic || fallback?.pic;
        if (pic) {
          anc['xdr:pic'] = pic;
          delete anc['mc:AlternateContent'];
        }
        return anc;
      });
      const anchors = Array.isArray(anchorsRaw) ? anchorsRaw : [anchorsRaw];

      for (const anchor of anchors) {
        const imageInfo = this.extractImageFromAnchor(anchor);
        if (imageInfo) {
          floatingImages.push(imageInfo);
        }
      }
    } catch (error) {
      console.error('解析drawing.xml失败:', error);
    }

    return floatingImages;
  }

  /**
   * 从锚点中提取图片信息
   */
  private extractImageFromAnchor(anchor: any): FloatingImageInfo | null {
    const pic = anchor?.['xdr:pic'] || anchor?.pic;
    if (!pic) return null;

    // 提取图片基本信息
    const nvPicPr = pic['xdr:nvPicPr'] || pic['nvPicPr'];
    const cNvPr = nvPicPr?.['xdr:cNvPr'] || nvPicPr?.['cNvPr'] || nvPicPr?.cNvPr;
    const blipFill = pic['xdr:blipFill'] || pic['blipFill'];
    const aBlip = blipFill?.['a:blip'] || blipFill?.blip;
    
    const description = cNvPr?.descr ?? '';
    const relationshipId = aBlip?.['r:embed'] || aBlip?.embed;
    
    if (!relationshipId) return null;

    // 提取位置信息
    const spPr = pic['xdr:spPr'] || pic['spPr'];
    const xfrm = spPr?.['a:xfrm'] || spPr?.xfrm;
    const off = xfrm?.['a:off'] || xfrm?.off;
    const ext = xfrm?.['a:ext'] || xfrm?.ext;
    
    const position = {
      x: Number(off?.x ?? 0),
      y: Number(off?.y ?? 0),
      width: Number(ext?.cx ?? 0),
      height: Number(ext?.cy ?? 0)
    };

    // 提取锚点信息
    const from = anchor?.['xdr:from'] || anchor?.from;
    const colIdx = Number(from?.['xdr:col'] ?? from?.col);
    const rowIdx = Number(from?.['xdr:row'] ?? from?.row);
    
    if (!Number.isFinite(rowIdx) || !Number.isFinite(colIdx)) return null;

    const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    
    // 生成更稳健的图片ID：忽略空/空白名称，优先 cNvPr.name，其次 cNvPr.id，再次关系ID，最后基于锚点生成
    const nameCandidate = typeof cNvPr?.name === 'string' ? cNvPr.name.trim() : '';
    const idCandidate = typeof cNvPr?.id === 'string' || typeof cNvPr?.id === 'number' ? String(cNvPr.id).trim() : '';
    const relCandidate = typeof relationshipId === 'string' ? relationshipId.trim() : '';
    const id = (nameCandidate || idCandidate || relCandidate || `floating_${cellRef}`);

    return {
      id,
      description,
      relationshipId,
      position,
      cellRef,
      rowIndex: rowIdx,
      colIndex: colIdx
    };
  }
}

/**
 * 浮动图片信息
 */
export interface FloatingImageInfo {
  /** 图片ID */
  id: string;
  /** 图片描述 */
  description: string;
  /** 关系ID */
  relationshipId: string;
  /** 位置信息 */
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  /** 锚点单元格引用 */
  cellRef: string;
  /** 行索引 */
  rowIndex: number;
  /** 列索引 */
  colIndex: number;
}
