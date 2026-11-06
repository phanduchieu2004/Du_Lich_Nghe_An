import { NestExpressApplication } from '@nestjs/platform-express';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export function configureViews(app: NestExpressApplication): void {
  const viewsDirectory = join(dirname(fileURLToPath(import.meta.url)), 'views');
  app.setBaseViewsDir(viewsDirectory);
  app.setViewEngine('ejs');
}
