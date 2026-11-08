import { Module } from '@nestjs/common';
import { TestController } from './test.controller.js';
import { MediaUploadService } from './media-upload.service.js';
import { TestService } from './test.service.js';

@Module({
  controllers: [TestController],
  providers: [MediaUploadService, TestService],
})
export class TestModule {}
