import { Module } from '@nestjs/common';
import { ReplacementController } from './replacement.controller';
import { ReplacementService } from './replacement.service';

@Module({
  controllers: [ReplacementController],
  providers: [ReplacementService],
  exports: [ReplacementService],
})
export class ReplacementModule {}

