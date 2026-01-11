"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ReplacementService = void 0;
const fs = require("fs");
const path = require("path");
const common_1 = require("@nestjs/common");
const XLSX = require('xlsx');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
let ReplacementService = class ReplacementService {
    replacementData = {
        CAUDAL: 100.05,
    };
    templatePath = path.join(process.cwd(), 'files', 'template.docx');
    resultPath = path.join(process.cwd(), 'result');
    async onModuleInit() {
        if (!fs.existsSync(this.resultPath)) {
            fs.mkdirSync(this.resultPath, { recursive: true });
        }
        this.generateDocumentToResult();
    }
    generateDocumentToResult() {
        try {
            const buffer = this.replaceInDocument();
            const outputPath = path.join(this.resultPath, 'documento-procesado.docx');
            fs.writeFileSync(outputPath, buffer);
            console.log(`Documento generado exitosamente en: ${outputPath}`);
        }
        catch (error) {
            console.error('Error al generar el documento:', error);
        }
    }
    fixSplitPlaceholders(xmlContent) {
        let fixedXml = xmlContent;
        const keys = Object.keys(this.replacementData);
        for (const key of keys) {
            const placeholder = `{{${key}}}`;
            if (fixedXml.includes(placeholder)) {
                continue;
            }
            const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const pattern = new RegExp(`(\\{\\{[^}]*${escapedKey}[^}]*\\}\\})|(\\{\\{[^}]*${escapedKey})|(${escapedKey}[^}]*\\}\\})`, 'gi');
            fixedXml = fixedXml.replace(pattern, (match) => {
                if (match.includes('{{') && match.includes('}}')) {
                    const keyPart = match.replace(/[{}]/g, '');
                    if (keyPart.includes(key) || key.includes(keyPart)) {
                        return placeholder;
                    }
                }
                return match;
            });
        }
        for (const key of keys) {
            const placeholder = `{{${key}}}`;
            const startPattern = `{{${key.substring(0, Math.ceil(key.length / 2))}`;
            const endPattern = `${key.substring(Math.ceil(key.length / 2))}}}`;
            if (fixedXml.includes(startPattern) &&
                fixedXml.includes(endPattern) &&
                !fixedXml.includes(placeholder)) {
                fixedXml = fixedXml.replace(new RegExp(startPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), placeholder);
                fixedXml = fixedXml.replace(new RegExp(endPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '');
            }
        }
        return fixedXml;
    }
    replaceInDocument() {
        const content = fs.readFileSync(this.templatePath, 'binary');
        const zip = new PizZip(content);
        const documentXml = zip.files['word/document.xml'];
        if (documentXml) {
            const xmlContent = documentXml.asText();
            const fixedXml = this.fixSplitPlaceholders(xmlContent);
            zip.file('word/document.xml', fixedXml);
        }
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: {
                start: '{{',
                end: '}}',
            },
            nullGetter: () => {
                return '';
            },
        });
        try {
            doc.render(this.replacementData);
        }
        catch (error) {
            if (error.properties && error.properties.errors instanceof Array) {
                const errorMessages = error.properties.errors
                    .map((e) => {
                    return e.properties
                        ? `${e.properties.explanation} (${e.properties.name})`
                        : e.message;
                })
                    .join('\n');
                throw new Error(`Error al procesar el template:\n${errorMessages}\n\nSugerencia: El placeholder {{CAUDAL}} está dividido en el documento Word. Intenta reescribirlo completo en una sola línea o usar un formato diferente.`);
            }
            throw error;
        }
        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });
        return buf;
    }
    replaceData(data) {
        const result = { ...data };
        for (const [key, value] of Object.entries(this.replacementData)) {
            if (key in result) {
                result[key] = value;
            }
        }
        return result;
    }
    getReplacementData() {
        return { ...this.replacementData };
    }
    updateReplacementData(newData) {
        Object.assign(this.replacementData, newData);
    }
    setReplacementValue(key, value) {
        this.replacementData[key] = value;
    }
    parseExcelToObject(fileBuffer) {
        try {
            const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
            });
            const dataRows = jsonData.slice(1);
            const result = {};
            for (const row of dataRows) {
                if (row && row.length >= 3) {
                    const key = row[0];
                    const value = row[2];
                    if (key !== undefined && key !== null) {
                        result[String(key)] = value;
                    }
                }
            }
            return result;
        }
        catch (error) {
            throw new Error(`Error al parsear el archivo Excel: ${error instanceof Error ? error.message : 'Error desconocido'}`);
        }
    }
};
exports.ReplacementService = ReplacementService;
exports.ReplacementService = ReplacementService = __decorate([
    (0, common_1.Injectable)()
], ReplacementService);
//# sourceMappingURL=replacement.service.js.map