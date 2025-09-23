import { XMLParser } from 'fast-xml-parser';

/**
 * 关系文件解析器
 * 负责解析Excel中的各种关系文件（.rels）
 */
export class RelationshipParser {
  private parser: XMLParser;

  constructor() {
    this.parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  }

  /**
   * 解析关系文件
   */
  parseRelationships(relsXml: string): Map<string, RelationshipInfo> {
    const relationships = new Map<string, RelationshipInfo>();
    
    try {
      const doc = this.parser.parse(relsXml);
      const rels = doc?.Relationships?.Relationship;
      if (!rels) return relationships;
      
      const list = Array.isArray(rels) ? rels : [rels];
      for (const rel of list) {
        const id = rel?.Id;
        const type = rel?.Type;
        const target = rel?.Target;
        const targetMode = rel?.TargetMode;
        if (id && type && target) {
          relationships.set(id, { id, type, target, targetMode });
        }
      }
    } catch (error) {
      console.error('解析关系文件失败:', error);
    }
    
    return relationships;
  }

  /**
   * 映射工作表名称到路径
   */
  mapSheetNamesToPaths(workbookXml: string, workbookRelsXml: string): Map<string, string> {
    const map = new Map<string, string>();
    
    try {
      const workbookDoc = this.parser.parse(workbookXml);
      const rels = this.parseRelationships(workbookRelsXml);
      
      const sheets = workbookDoc?.workbook?.sheets?.sheet || workbookDoc?.['x:workbook']?.sheets?.sheet || [];
      const list = Array.isArray(sheets) ? sheets : [sheets];
      
      for (const sheet of list) {
        const name = sheet?.name;
        const rid = sheet?.['r:id'] || sheet?.id;
        
        if (!name || !rid) continue;
        
        const rel = rels.get(rid);
        if (!rel) continue;
        
        // Target like 'worksheets/sheet1.xml'
        map.set(name, rel.target.replace(/^\.\//, 'worksheets/'));
      }
    } catch (error) {
      console.error('映射工作表名称失败:', error);
    }
    
    return map;
  }

  /**
   * 标准化目标路径
   */
  normalizeTargetPath(fromFolder: 'worksheets' | 'drawings', target: string): string {
    // Targets in rels are relative to their folder, e.g. '../media/image1.jpeg' or '../drawings/drawing1.xml'
    if (fromFolder === 'worksheets') {
      // base: xl/worksheets/ -> normalize to path under xl/
      if (target.startsWith('../')) return target.replace(/^\.\.\//, '');
      if (target.startsWith('/')) return target.replace(/^\/+/, '');
      if (!target.startsWith('worksheets/') && !target.startsWith('drawings/') && !target.startsWith('media/')) {
        return `worksheets/${target}`;
      }
      return target;
    } else {
      // from drawings
      if (target.startsWith('../')) return target.replace(/^\.\.\//, '');
      if (target.startsWith('/')) return target.replace(/^\/+/, '');
      if (!target.startsWith('drawings/') && !target.startsWith('media/')) {
        return `drawings/${target}`;
      }
      return target;
    }
  }

  /**
   * 查找drawing关系
   */
  findDrawingRelationship(relationships: Map<string, RelationshipInfo>): RelationshipInfo | null {
    return Array.from(relationships.values()).find(r => 
      r.type.includes('/relationships/drawing')
    ) || null;
  }
}

/**
 * 关系信息
 */
export interface RelationshipInfo {
  /** 关系ID */
  id: string;
  /** 关系类型 */
  type: string;
  /** 目标路径 */
  target: string;
  /** 目标模式（External 表示外链） */
  targetMode?: string;
}
