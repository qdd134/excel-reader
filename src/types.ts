/**
 * Excel单元格图片信息
 */
export interface CellImage {
  /** 图片ID，对应cellimages.xml中的name属性 */
  id: string;
  /** 图片描述，对应cellimages.xml中的descr属性 */
  description: string;
  /** 图片的base64编码数据 */
  base64: string;
  /** 图片MIME类型 */
  mimeType: string;
  /** 图片在单元格中的位置信息 */
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  /** 关联的关系ID，用于查找实际图片文件 */
  relationshipId: string;
}

/**
 * Excel单元格数据
 */
export interface CellData {
  /** 单元格引用（如A1, B2等） */
  ref: string;
  /** 单元格值 */
  value: string | number;
  /** 单元格类型 */
  type: 'string' | 'number' | 'formula' | 'image';
  /** 单元格样式ID */
  styleId?: number;
  /** 如果是图片单元格，包含图片信息 */
  image?: CellImage;
  /** 如果是公式单元格，包含公式内容 */
  formula?: string;
}

/**
 * Excel行数据
 */
export interface RowData {
  /** 行号 */
  rowNumber: number;
  /** 行高 */
  height?: number;
  /** 是否自定义行高 */
  customHeight?: boolean;
  /** 单元格数据数组 */
  cells: CellData[];
  /** 该行包含的图片数量 */
  imageCount: number;
  /** 该行包含图片的单元格引用 */
  imageCells: string[];
}

/**
 * Excel工作表数据
 */
export interface WorksheetData {
  /** 工作表名称 */
  name: string;
  /** 工作表数据范围 */
  dimension: {
    start: string;
    end: string;
  };
  /** 行数据数组 */
  rows: RowData[];
  /** 列配置 */
  columns: {
    min: number;
    max: number;
    width: number;
    customWidth: boolean;
  }[];
  /** 该工作表包含的图片总数 */
  totalImages: number;
  /** 包含图片的行数 */
  rowsWithImages: number;
}

/**
 * Excel文件解析结果
 */
export interface ExcelParseResult {
  /** 工作表数据数组 */
  worksheets: WorksheetData[];
  /** 所有图片的映射表，key为图片ID */
  images: Map<string, CellImage>;
  /** 解析过程中的错误信息 */
  errors: string[];
}

/**
 * 解析选项
 */
export interface ParseOptions {
  /** 是否包含图片数据 */
  includeImages?: boolean;
  /** 图片质量（0-1，仅对JPEG有效） */
  imageQuality?: number;
  /** 是否包含空行 */
  includeEmptyRows?: boolean;
  /** 是否包含空列 */
  includeEmptyColumns?: boolean;
}

/**
 * 图片提取结果
 */
export interface ImageExtractionResult {
  /** 图片ID */
  id: string;
  /** 图片描述 */
  description: string;
  /** 图片的base64数据 */
  base64: string;
  /** 图片MIME类型 */
  mimeType: string;
  /** 原始图片数据 */
  rawData: Uint8Array;
}
