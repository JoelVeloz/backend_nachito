import { Response } from 'express';
import { ReplacementService } from './replacement.service';
export declare class ReplacementController {
    private readonly replacementService;
    constructor(replacementService: ReplacementService);
    generateDocument(res: Response): void;
    updateAndGenerate(data: Record<string, unknown>): void;
    uploadExcel(file: Express.Multer.File): Record<string, unknown>;
    processExcelAndDownload(file: Express.Multer.File, res: Response): void;
}
