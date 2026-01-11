import {
  Controller,
  Get,
  Post,
  Body,
  Res,
  Header,
  UseInterceptors,
  UploadedFile,
  BadRequestException,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { ReplacementService } from './replacement.service';

@Controller('replacement')
export class ReplacementController {
  constructor(private readonly replacementService: ReplacementService) {}

  /**
   * Genera el documento Word con las variables reemplazadas
   * GET /replacement
   */
  @Get()
  @Header(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  )
  @Header(
    'Content-Disposition',
    'attachment; filename="documento-procesado.docx"',
  )
  generateDocument(@Res() res: Response): void {
    const buffer = this.replacementService.replaceInDocument();
    res.send(buffer);
  }

  /**
   * Actualiza los valores de reemplazo y regenera el documento
   * POST /replacement
   */
  @Post()
  updateAndGenerate(@Body() data: Record<string, unknown>): void {
    this.replacementService.updateReplacementData(data);
    this.replacementService.generateDocumentToResult();
  }

  /**
   * Parsea un archivo Excel y retorna un objeto con CODE como clave y value como valor
   * POST /replacement/upload-excel
   */
  @Post('upload-excel')
  @UseInterceptors(FileInterceptor('file'))
  uploadExcel(
    @UploadedFile() file: Express.Multer.File,
  ): Record<string, unknown> {
    if (!file) {
      throw new BadRequestException('No se proporcionó ningún archivo');
    }

    // Verifica que sea un archivo Excel
    const allowedMimeTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/vnd.ms-excel.sheet.macroEnabled.12',
    ];

    if (!allowedMimeTypes.includes(file.mimetype)) {
      throw new BadRequestException(
        'El archivo debe ser un Excel (.xlsx, .xls)',
      );
    }

    try {
      const result = this.replacementService.parseExcelToObject(file.buffer);
      return result;
    } catch (error) {
      throw new BadRequestException(
        error instanceof Error
          ? error.message
          : 'Error al procesar el archivo Excel',
      );
    }
  }

  /**
   * Sube un archivo Excel, genera el JSON y descarga el documento Word con los reemplazos aplicados
   * POST /replacement/process-excel
   */
  @Post('process-excel')
  @UseInterceptors(FileInterceptor('file'))
  processExcelAndDownload(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ): void {
    if (!file) {
      throw new BadRequestException('No se proporcionó ningún archivo');
    }

    // Verifica que sea un archivo Excel
    const allowedMimeTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/vnd.ms-excel.sheet.macroEnabled.12',
    ];

    if (!allowedMimeTypes.includes(file.mimetype)) {
      throw new BadRequestException(
        'El archivo debe ser un Excel (.xlsx, .xls)',
      );
    }

    try {
      // Parsea el Excel a JSON
      const jsonData = this.replacementService.parseExcelToObject(file.buffer);

      // Actualiza los datos de reemplazo con el JSON generado
      this.replacementService.updateReplacementData(jsonData);

      // Genera el documento Word con los reemplazos aplicados
      const buffer = this.replacementService.replaceInDocument();

      // Establece los headers para la descarga del archivo Word
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      );
      res.setHeader(
        'Content-Disposition',
        'attachment; filename="documento-procesado.docx"',
      );

      // Envía el documento como descarga
      res.send(buffer);
    } catch (error) {
      throw new BadRequestException(
        error instanceof Error
          ? error.message
          : 'Error al procesar el archivo Excel y generar el documento',
      );
    }
  }
}
