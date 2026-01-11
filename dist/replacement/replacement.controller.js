"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ReplacementController = void 0;
const common_1 = require("@nestjs/common");
const platform_express_1 = require("@nestjs/platform-express");
const replacement_service_1 = require("./replacement.service");
let ReplacementController = class ReplacementController {
    replacementService;
    constructor(replacementService) {
        this.replacementService = replacementService;
    }
    generateDocument(res) {
        const buffer = this.replacementService.replaceInDocument();
        res.send(buffer);
    }
    updateAndGenerate(data) {
        this.replacementService.updateReplacementData(data);
        this.replacementService.generateDocumentToResult();
    }
    uploadExcel(file) {
        if (!file) {
            throw new common_1.BadRequestException('No se proporcionó ningún archivo');
        }
        const allowedMimeTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'application/vnd.ms-excel.sheet.macroEnabled.12',
        ];
        if (!allowedMimeTypes.includes(file.mimetype)) {
            throw new common_1.BadRequestException('El archivo debe ser un Excel (.xlsx, .xls)');
        }
        try {
            const result = this.replacementService.parseExcelToObject(file.buffer);
            return result;
        }
        catch (error) {
            throw new common_1.BadRequestException(error instanceof Error
                ? error.message
                : 'Error al procesar el archivo Excel');
        }
    }
    processExcelAndDownload(file, res) {
        if (!file) {
            throw new common_1.BadRequestException('No se proporcionó ningún archivo');
        }
        const allowedMimeTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'application/vnd.ms-excel.sheet.macroEnabled.12',
        ];
        if (!allowedMimeTypes.includes(file.mimetype)) {
            throw new common_1.BadRequestException('El archivo debe ser un Excel (.xlsx, .xls)');
        }
        try {
            const jsonData = this.replacementService.parseExcelToObject(file.buffer);
            this.replacementService.updateReplacementData(jsonData);
            const buffer = this.replacementService.replaceInDocument();
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename="documento-procesado.docx"');
            res.send(buffer);
        }
        catch (error) {
            throw new common_1.BadRequestException(error instanceof Error
                ? error.message
                : 'Error al procesar el archivo Excel y generar el documento');
        }
    }
};
exports.ReplacementController = ReplacementController;
__decorate([
    (0, common_1.Get)(),
    (0, common_1.Header)('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    (0, common_1.Header)('Content-Disposition', 'attachment; filename="documento-procesado.docx"'),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", void 0)
], ReplacementController.prototype, "generateDocument", null);
__decorate([
    (0, common_1.Post)(),
    __param(0, (0, common_1.Body)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", void 0)
], ReplacementController.prototype, "updateAndGenerate", null);
__decorate([
    (0, common_1.Post)('upload-excel'),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Object)
], ReplacementController.prototype, "uploadExcel", null);
__decorate([
    (0, common_1.Post)('process-excel'),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", void 0)
], ReplacementController.prototype, "processExcelAndDownload", null);
exports.ReplacementController = ReplacementController = __decorate([
    (0, common_1.Controller)('replacement'),
    __metadata("design:paramtypes", [replacement_service_1.ReplacementService])
], ReplacementController);
//# sourceMappingURL=replacement.controller.js.map