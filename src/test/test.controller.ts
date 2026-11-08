import {
  BadRequestException,
  Body,
  Controller,
  Delete,
  Get,
  NotFoundException,
  Param,
  Patch,
  Post,
  Query,
  Res,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Response } from 'express';
import { MediaUploadService } from './media-upload.service.js';
import { TestService } from './test.service.js';

interface UploadedExcelFile {
  originalname: string;
  buffer: Buffer;
}

@Controller('test')
export class TestController {
  constructor(
    private readonly testService: TestService,
    private readonly mediaUploadService: MediaUploadService,
  ) {}

  @Get()
  showPage(@Res() response: Response): void {
    if (process.env.NODE_ENV === 'production') {
      throw new NotFoundException();
    }
    response.sendFile(
      join(dirname(fileURLToPath(import.meta.url)), 'index.html'),
    );
  }

  @Get('media')
  showMediaPage(@Res() response: Response): void {
    if (process.env.NODE_ENV === 'production') {
      throw new NotFoundException();
    }
    response.sendFile(
      join(dirname(fileURLToPath(import.meta.url)), 'media.html'),
    );
  }

  @Get('api/media')
  getMediaLibrary() {
    return this.testService.getMediaLibrary();
  }

  @Get('api/tables')
  getTables() {
    return this.testService.getTables();
  }

  @Get('api/tables/:table')
  getRows(
    @Param('table') table: string,
    @Query('q') search = '',
    @Query('page') page = '1',
  ) {
    const pageNumber = Number(page);
    if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) {
      throw new BadRequestException('Trang phải là số nguyên dương.');
    }

    return this.testService.getRows(table, search, pageNumber);
  }

  @Post('api/tables/:table')
  createRow(@Param('table') table: string, @Body() body: unknown) {
    return this.testService.createRow(table, body);
  }

  @Patch('api/tables/:table')
  updateRow(@Param('table') table: string, @Body() body: unknown) {
    return this.testService.updateRow(table, body);
  }

  @Delete('api/tables/:table')
  deleteRow(@Param('table') table: string, @Body() body: unknown) {
    return this.testService.deleteRow(table, body);
  }

  @Post('api/import')
  @UseInterceptors(
    FileInterceptor('file', {
      limits: { fileSize: 20 * 1024 * 1024, files: 1 },
    }),
  )
  importExcel(@UploadedFile() file: UploadedExcelFile | undefined) {
    return this.testService.importExcel(this.getExcelBuffer(file));
  }

  @Post('api/upload')
  @UseInterceptors(
    FileInterceptor('file', {
      limits: { fileSize: 100 * 1024 * 1024, files: 1 },
    }),
  )
  uploadMedia(
    @UploadedFile()
    file: { buffer: Buffer; mimetype: string } | undefined,
  ) {
    return this.mediaUploadService.upload(file);
  }

  private getExcelBuffer(file: UploadedExcelFile | undefined): Buffer {
    if (!file) {
      throw new BadRequestException('Hãy chọn một file Excel .xlsx để nhập.');
    }
    if (!file.originalname.toLocaleLowerCase('en-US').endsWith('.xlsx')) {
      throw new BadRequestException(
        'Chỉ chấp nhận file Excel định dạng .xlsx.',
      );
    }
    return file.buffer;
  }
}
