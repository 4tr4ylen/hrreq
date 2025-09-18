import { IRequestFormData } from '../models/IRequest';

export interface IValidationResult {
  isValid: boolean;
  errors: string[];
}

export class ValidationUtils {
  // Validate request form data
  public static validateRequestForm(data: IRequestFormData): IValidationResult {
    const errors: string[] = [];

    // Title validation
    if (!data.title || data.title.trim().length === 0) {
      errors.push('Title is required');
    } else if (data.title.trim().length > 255) {
      errors.push('Title must be 255 characters or less');
    }

    // Request type validation
    if (!data.requestType || data.requestType.trim().length === 0) {
      errors.push('Request type is required');
    }

    // Description validation
    if (!data.description || data.description.trim().length === 0) {
      errors.push('Description is required');
    } else if (data.description.trim().length > 2000) {
      errors.push('Description must be 2000 characters or less');
    }

    // Department validation
    if (!data.department || data.department.trim().length === 0) {
      errors.push('Department is required');
    }

    // Attachment validation
    if (!data.attachment) {
      errors.push('Exactly one attachment is required');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  // Validate file upload
  public static validateFile(file: File, maxSizeMB: number = 10, allowedTypes: string[] = []): IValidationResult {
    const errors: string[] = [];

    // File size validation
    const maxSizeBytes = maxSizeMB * 1024 * 1024;
    if (file.size > maxSizeBytes) {
      errors.push(`File size must be ${maxSizeMB}MB or less`);
    }

    // File type validation
    if (allowedTypes.length > 0) {
      const fileExtension = file.name.split('.').pop()?.toLowerCase();
      const mimeType = file.type.toLowerCase();
      
      const isValidType = allowedTypes.some(type => {
        const typeLower = type.toLowerCase();
        return fileExtension === typeLower || mimeType.includes(typeLower);
      });

      if (!isValidType) {
        errors.push(`File type not allowed. Allowed types: ${allowedTypes.join(', ')}`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  // Get allowed file types
  public static getAllowedFileTypes(): string[] {
    return [
      'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
      'txt', 'rtf', 'jpg', 'jpeg', 'png', 'gif', 'bmp',
      'zip', 'rar', '7z'
    ];
  }

  // Format file size
  public static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  // Validate email format
  public static isValidEmail(email: string): boolean {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  // Sanitize HTML content
  public static sanitizeHtml(html: string): string {
    const div = document.createElement('div');
    div.textContent = html;
    return div.innerHTML;
  }

  // Validate required fields
  public static validateRequired(value: any, fieldName: string): string | null {
    if (!value || (typeof value === 'string' && value.trim().length === 0)) {
      return `${fieldName} is required`;
    }
    return null;
  }

  // Validate string length
  public static validateStringLength(value: string, fieldName: string, maxLength: number): string | null {
    if (value && value.length > maxLength) {
      return `${fieldName} must be ${maxLength} characters or less`;
    }
    return null;
  }
} 