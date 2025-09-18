import { ValidationUtils } from './validation';

export interface IFileInfo {
  name: string;
  size: number;
  type: string;
  lastModified: number;
}

export class AttachmentUtils {
  // Convert File to ArrayBuffer
  public static async fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          resolve(reader.result);
        } else {
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  // Get file info
  public static getFileInfo(file: File): IFileInfo {
    return {
      name: file.name,
      size: file.size,
      type: file.type,
      lastModified: file.lastModified
    };
  }

  // Validate single file upload
  public static validateSingleFile(file: File | null, maxSizeMB: number = 10): string[] {
    const errors: string[] = [];

    if (!file) {
      errors.push('Please select a file to upload');
      return errors;
    }

    // Validate file size
    const maxSizeBytes = maxSizeMB * 1024 * 1024;
    if (file.size > maxSizeBytes) {
      errors.push(`File size must be ${maxSizeMB}MB or less`);
    }

    // Validate file type
    const allowedTypes = ValidationUtils.getAllowedFileTypes();
    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    const mimeType = file.type.toLowerCase();
    
    const isValidType = allowedTypes.some(type => {
      const typeLower = type.toLowerCase();
      return fileExtension === typeLower || mimeType.includes(typeLower);
    });

    if (!isValidType) {
      errors.push(`File type not allowed. Allowed types: ${allowedTypes.join(', ')}`);
    }

    return errors;
  }

  // Format file size for display
  public static formatFileSize(bytes: number): string {
    return ValidationUtils.formatFileSize(bytes);
  }

  // Get file icon based on type
  public static getFileIcon(fileName: string): string {
    const extension = fileName.split('.').pop()?.toLowerCase();
    
    switch (extension) {
      case 'pdf':
        return 'PDF';
      case 'doc':
      case 'docx':
        return 'Word';
      case 'xls':
      case 'xlsx':
        return 'Excel';
      case 'ppt':
      case 'pptx':
        return 'PowerPoint';
      case 'txt':
        return 'Text';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
      case 'bmp':
        return 'Image';
      case 'zip':
      case 'rar':
      case '7z':
        return 'Archive';
      default:
        return 'Document';
    }
  }

  // Get file color based on type
  public static getFileColor(fileName: string): string {
    const extension = fileName.split('.').pop()?.toLowerCase();
    
    switch (extension) {
      case 'pdf':
        return '#d13438'; // Red
      case 'doc':
      case 'docx':
        return '#0078d4'; // Blue
      case 'xls':
      case 'xlsx':
        return '#107c10'; // Green
      case 'ppt':
      case 'pptx':
        return '#d83b01'; // Orange
      case 'txt':
        return '#605e5c'; // Gray
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
      case 'bmp':
        return '#881798'; // Purple
      case 'zip':
      case 'rar':
      case '7z':
        return '#ff8c00'; // Dark Orange
      default:
        return '#605e5c'; // Gray
    }
  }

  // Create download link for attachment
  public static createDownloadLink(serverRelativeUrl: string): string {
    const baseUrl = window.location.origin;
    return `${baseUrl}${serverRelativeUrl}`;
  }

  // Check if file is image
  public static isImage(file: File): boolean {
    return file.type.startsWith('image/');
  }

  // Check if file is document
  public static isDocument(file: File): boolean {
    const documentTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'text/plain'
    ];
    return documentTypes.includes(file.type);
  }

  // Get file preview URL (for images)
  public static createPreviewUrl(file: File): string {
    if (this.isImage(file)) {
      return URL.createObjectURL(file);
    }
    return '';
  }

  // Clean up preview URL
  public static revokePreviewUrl(url: string): void {
    if (url && url.startsWith('blob:')) {
      URL.revokeObjectURL(url);
    }
  }
} 