import {
  BadGatewayException,
  BadRequestException,
  Injectable,
  NotFoundException,
  ServiceUnavailableException,
} from '@nestjs/common';
import { v2 as cloudinary } from 'cloudinary';

export interface UploadedMediaFile {
  buffer: Buffer;
  mimetype: string;
}

@Injectable()
export class MediaUploadService {
  async upload(
    file: UploadedMediaFile | undefined,
  ): Promise<{ secureUrl: string }> {
    if (process.env.NODE_ENV === 'production') {
      throw new NotFoundException();
    }
    if (!file || file.buffer.length === 0) {
      throw new BadRequestException('Hãy chọn một tệp để tải lên.');
    }
    if (!/^(image|audio|video)\//i.test(file.mimetype)) {
      throw new BadRequestException(
        'Chỉ chấp nhận tệp hình ảnh, âm thanh hoặc video.',
      );
    }
    const mediaType = file.mimetype.split('/')[0].toLocaleLowerCase('en-US');
    const mediaFolder =
      mediaType === 'image'
        ? 'hinh-anh'
        : mediaType === 'audio'
          ? 'am-thanh'
          : 'video';

    const cloudName = process.env.CLOUDINARY_CLOUD_NAME;
    const apiKey = process.env.CLOUDINARY_API_KEY;
    const apiSecret = process.env.CLOUDINARY_API_SECRET;
    if (!cloudName || !apiKey || !apiSecret) {
      throw new ServiceUnavailableException(
        'Chưa cấu hình Cloudinary. Hãy đặt CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY và CLOUDINARY_API_SECRET trong .env.',
      );
    }

    cloudinary.config({
      cloud_name: cloudName,
      api_key: apiKey,
      api_secret: apiSecret,
      secure: true,
    });

    return new Promise((resolve, reject) => {
      const uploadStream = cloudinary.uploader.upload_stream(
        {
          folder: `du-lich-nghe-an/thuyet-minh/${mediaFolder}`,
          resource_type: 'auto',
          use_filename: true,
          unique_filename: true,
        },
        (error, result) => {
          if (error) {
            reject(
              new BadGatewayException(
                `Không thể tải tệp lên Cloudinary: ${error.message}`,
              ),
            );
            return;
          }
          if (!result?.secure_url) {
            reject(
              new BadGatewayException(
                'Cloudinary không trả về đường dẫn bảo mật của tệp.',
              ),
            );
            return;
          }
          resolve({ secureUrl: result.secure_url });
        },
      );
      uploadStream.end(file.buffer);
    });
  }
}
