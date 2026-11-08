import {
  BadRequestException,
  ServiceUnavailableException,
} from '@nestjs/common';
import { MediaUploadService } from './media-upload.service.js';

describe('MediaUploadService', () => {
  const service = new MediaUploadService();

  beforeEach(() => {
    vi.stubEnv('NODE_ENV', 'test');
    vi.stubEnv('CLOUDINARY_CLOUD_NAME', '');
    vi.stubEnv('CLOUDINARY_API_KEY', '');
    vi.stubEnv('CLOUDINARY_API_SECRET', '');
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it('rejects files that are not images, audio, or video', async () => {
    await expect(
      service.upload({
        buffer: Buffer.from('document'),
        mimetype: 'application/pdf',
      }),
    ).rejects.toBeInstanceOf(BadRequestException);
  });

  it('reports missing Cloudinary credentials without attempting an upload', async () => {
    await expect(
      service.upload({
        buffer: Buffer.from('image'),
        mimetype: 'image/png',
      }),
    ).rejects.toBeInstanceOf(ServiceUnavailableException);
  });
});
