import * as JSZip from 'jszip';
import { CellImage, ImageExtractionResult } from './types';

/**
 * 图片数据提取器
 * 负责从Excel文件中提取图片的二进制数据并转换为base64
 */
export class ImageExtractor {
  private zip: JSZip | null = null;

  /**
   * 设置JSZip实例
   */
  setZip(zip: JSZip): void {
    this.zip = zip;
  }

  /**
   * 提取图片数据
   */
  async extractImageData(imagePath: string): Promise<ImageExtractionResult | null> {
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
      console.error('提取图片数据失败:', error);
      return null;
    }
  }

  /**
   * 从图片路径提取多个图片数据
   */
  async extractMultipleImages(imagePaths: string[]): Promise<Map<string, ImageExtractionResult>> {
    const results = new Map<string, ImageExtractionResult>();
    
    const promises = imagePaths.map(async (path) => {
      const result = await this.extractImageData(path);
      if (result) {
        results.set(path, result);
      }
    });

    await Promise.all(promises);
    return results;
  }

  /**
   * 创建CellImage对象
   */
  createCellImage(
    id: string,
    description: string,
    relationshipId: string,
    position: { x: number; y: number; width: number; height: number },
    imageResult: ImageExtractionResult
  ): CellImage {
    return {
      id,
      description,
      base64: imageResult.base64,
      mimeType: imageResult.mimeType,
      position,
      relationshipId
    };
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
   * 检查图片文件是否存在
   */
  async checkImageExists(imagePath: string): Promise<boolean> {
    if (!this.zip) return false;
    
    try {
      const imageFile = this.zip.file(`xl/${imagePath}`);
      return imageFile !== null;
    } catch {
      return false;
    }
  }

  /**
   * 获取图片文件大小
   */
  async getImageSize(imagePath: string): Promise<number | null> {
    if (!this.zip) return null;
    
    try {
      const imageFile = this.zip.file(`xl/${imagePath}`);
      if (!imageFile) return null;
      
      const buffer = await imageFile.async('nodebuffer');
      return buffer.length;
    } catch {
      return null;
    }
  }
}
