import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ReplacementModule } from './replacement/replacement.module';

@Module({
  imports: [ReplacementModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
