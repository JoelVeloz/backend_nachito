import { OnModuleInit } from '@nestjs/common';
export declare class ReplacementService implements OnModuleInit {
    private readonly replacementData;
    private readonly templatePath;
    private readonly resultPath;
    onModuleInit(): Promise<void>;
    generateDocumentToResult(): void;
    private fixSplitPlaceholders;
    replaceInDocument(): Buffer;
    replaceData(data: Record<string, unknown>): Record<string, unknown>;
    getReplacementData(): Record<string, unknown>;
    updateReplacementData(newData: Record<string, unknown>): void;
    setReplacementValue(key: string, value: unknown): void;
    parseExcelToObject(fileBuffer: Buffer): Record<string, unknown>;
}
