import * as fs from 'fs';
import * as path from 'path';

import { Injectable, OnModuleInit } from '@nestjs/common';

// eslint-disable-next-line @typescript-eslint/no-require-imports
const XLSX = require('xlsx');

// eslint-disable-next-line @typescript-eslint/no-require-imports
const Docxtemplater = require('docxtemplater');
// eslint-disable-next-line @typescript-eslint/no-require-imports
const PizZip = require('pizzip');

@Injectable()
export class ReplacementService implements OnModuleInit {
  // JSON de configuración con los datos de reemplazo
  private readonly replacementData: Record<string, unknown> = {
    CAUDAL: 100.05,
  };

  private readonly templatePath = path.join(
    process.cwd(),
    'files',
    'template.docx',
  );

  private readonly resultPath = path.join(process.cwd(), 'result');

  /**
   * Se ejecuta cuando el módulo se inicializa
   */
  async onModuleInit(): Promise<void> {
    // Crea la carpeta result si no existe
    if (!fs.existsSync(this.resultPath)) {
      fs.mkdirSync(this.resultPath, { recursive: true });
    }

    // Genera el documento procesado
    this.generateDocumentToResult();
  }

  /**
   * Genera el documento procesado y lo guarda en la carpeta result
   */
  generateDocumentToResult(): void {
    try {
      const buffer = this.replaceInDocument();
      const outputPath = path.join(this.resultPath, 'documento-procesado.docx');
      fs.writeFileSync(outputPath, buffer);
      console.log(`Documento generado exitosamente en: ${outputPath}`);
    } catch (error) {
      console.error('Error al generar el documento:', error);
    }
  }

  /**
   * Pre-procesa el XML del documento para unir textos divididos que contengan placeholders
   * Word divide los textos en múltiples elementos <w:t>, esta función los une
   */
  private fixSplitPlaceholders(xmlContent: string): string {
    let fixedXml = xmlContent;

    // Lista de todas las keys del replacementData para buscar sus placeholders
    const keys = Object.keys(this.replacementData);

    for (const key of keys) {
      const placeholder = `{{${key}}}`;

      // Si el placeholder está completo, no hacer nada
      if (fixedXml.includes(placeholder)) {
        continue;
      }

      // Busca patrones donde el placeholder está dividido
      // Ejemplo: {{CAUD en un <w:t> y AL}} en otro
      // Usamos una expresión más flexible que busque el patrón dividido
      const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

      // Busca {{ seguido de cualquier parte de la key, posiblemente dividido
      // y luego }} posiblemente después de más texto
      const pattern = new RegExp(
        `(\\{\\{[^}]*${escapedKey}[^}]*\\}\\})|(\\{\\{[^}]*${escapedKey})|(${escapedKey}[^}]*\\}\\})`,
        'gi',
      );

      // Si encontramos el patrón dividido, intentamos unirlo
      // Nota: Esta es una solución simplificada. Para casos complejos,
      // sería mejor usar un parser XML real
      fixedXml = fixedXml.replace(pattern, (match) => {
        // Si el match contiene el placeholder completo o casi completo, lo reemplazamos
        if (match.includes('{{') && match.includes('}}')) {
          // Extrae la parte de la key que está presente
          const keyPart = match.replace(/[{}]/g, '');
          if (keyPart.includes(key) || key.includes(keyPart)) {
            return placeholder;
          }
        }
        return match;
      });
    }

    // Solución más robusta: busca y reemplaza directamente los placeholders conocidos
    // incluso si están parcialmente divididos
    for (const key of keys) {
      const placeholder = `{{${key}}}`;
      // Busca variaciones del placeholder que puedan estar divididas
      // Ejemplo: busca "{{CAUD" seguido eventualmente de "AL}}"
      const startPattern = `{{${key.substring(0, Math.ceil(key.length / 2))}`;
      const endPattern = `${key.substring(Math.ceil(key.length / 2))}}}`;

      if (
        fixedXml.includes(startPattern) &&
        fixedXml.includes(endPattern) &&
        !fixedXml.includes(placeholder)
      ) {
        // Reemplaza las partes divididas con el placeholder completo
        fixedXml = fixedXml.replace(
          new RegExp(startPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
          placeholder,
        );
        fixedXml = fixedXml.replace(
          new RegExp(endPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
          '',
        );
      }
    }

    return fixedXml;
  }

  /**
   * Reemplaza los valores en el documento Word template.docx según el JSON de configuración
   * @returns Buffer del documento Word procesado
   */
  replaceInDocument(): Buffer {
    // Lee el archivo template.docx
    const content = fs.readFileSync(this.templatePath, 'binary');

    // Crea una instancia de PizZip con el contenido del documento
    const zip = new PizZip(content);

    // Pre-procesa el XML del documento para unir textos divididos
    const documentXml = zip.files['word/document.xml'];
    if (documentXml) {
      const xmlContent = documentXml.asText();
      const fixedXml = this.fixSplitPlaceholders(xmlContent);
      zip.file('word/document.xml', fixedXml);
    }

    // Crea una instancia de Docxtemplater con configuración para manejar textos divididos
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: '{{',
        end: '}}',
      },
      // Maneja valores nulos o no encontrados
      nullGetter: () => {
        return '';
      },
    });

    try {
      // Reemplaza los valores en el documento usando el JSON de configuración
      doc.render(this.replacementData);
    } catch (error) {
      // Si hay errores de template, los muestra con más detalle
      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map((e) => {
            return e.properties
              ? `${e.properties.explanation} (${e.properties.name})`
              : e.message;
          })
          .join('\n');
        throw new Error(
          `Error al procesar el template:\n${errorMessages}\n\nSugerencia: El placeholder {{CAUDAL}} está dividido en el documento Word. Intenta reescribirlo completo en una sola línea o usar un formato diferente.`,
        );
      }
      throw error;
    }

    // Genera el documento procesado
    const buf = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    return buf;
  }

  /**
   * Reemplaza los valores en un JSON según el mapeo de configuración
   * @param data JSON de entrada que se desea procesar
   * @returns JSON con los valores reemplazados
   */
  replaceData(data: Record<string, unknown>): Record<string, unknown> {
    const result = { ...data };

    // Itera sobre las keys del JSON de configuración
    for (const [key, value] of Object.entries(this.replacementData)) {
      // Si la key existe en el JSON de entrada, reemplaza su valor
      if (key in result) {
        result[key] = value;
      }
    }

    return result;
  }

  /**
   * Obtiene el JSON de configuración de reemplazo
   * @returns JSON con los datos de reemplazo
   */
  getReplacementData(): Record<string, unknown> {
    return { ...this.replacementData };
  }

  /**
   * Actualiza el JSON de configuración de reemplazo
   * @param newData Nuevo JSON de configuración
   */
  updateReplacementData(newData: Record<string, unknown>): void {
    Object.assign(this.replacementData, newData);
  }

  /**
   * Reemplaza un valor específico en el JSON de configuración
   * @param key Key a reemplazar
   * @param value Nuevo valor
   */
  setReplacementValue(key: string, value: unknown): void {
    this.replacementData[key] = value;
  }

  /**
   * Parsea un archivo Excel y genera un objeto con la primera columna como clave y la tercera como valor
   * @param fileBuffer Buffer del archivo Excel
   * @returns Objeto con estructura { primeraColumna: terceraColumna }
   */
  parseExcelToObject(fileBuffer: Buffer): Record<string, unknown> {
    try {
      // Lee el archivo Excel desde el buffer
      const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

      // Obtiene la primera hoja del libro
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convierte la hoja a array de arrays (header: 1 significa que cada fila es un array)
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      }) as unknown[][];

      // Elimina la primera fila (cabecera)
      const dataRows = jsonData.slice(1);

      // Genera el objeto con la primera columna como clave y la tercera como valor
      const result: Record<string, unknown> = {};

      for (const row of dataRows) {
        // Verifica que la fila tenga al menos 3 columnas
        if (row && row.length >= 3) {
          const key = row[0]; // Primera columna
          const value = row[2]; // Tercera columna

          // Solo agrega si la clave existe y no es undefined/null
          if (key !== undefined && key !== null) {
            result[String(key)] = value;
          }
        }
      }

      return result;
    } catch (error) {
      throw new Error(
        `Error al parsear el archivo Excel: ${error instanceof Error ? error.message : 'Error desconocido'}`,
      );
    }
  }
}
