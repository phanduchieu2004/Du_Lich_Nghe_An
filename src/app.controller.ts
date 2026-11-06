import { Controller, Get, Render } from '@nestjs/common';

@Controller()
export class AppController {
  @Get()
  @Render('index')
  getIndex(): Record<string, never> {
    return {};
  }

  @Get('chuc-nang')
  @Render('trang-chuc-nang')
  getFunctionButtons(): { buttonIds: string[] } {
    return { buttonIds: ['quan-ly-co-so-du-lieu'] };
  }
}
