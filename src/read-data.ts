import * as fs from "fs";
import { Rows } from "./types";
import { Workbook } from "exceljs";

interface Input {
  filePath: string;
}

export async function readData(input: Input): Promise<Rows> {
  const isFileExists = fs.existsSync(input.filePath);

  if (!isFileExists) {
    console.log("📄 Arquivo não existe, será criado com os novos dados.");
    return [];
  }

  try {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(input.filePath);

    // Procura pela aba "SKU x DROP" primeiro, senão usa a primeira aba
    let worksheet = workbook.getWorksheet("SKU x DROP");
    if (!worksheet) {
      worksheet = workbook.getWorksheet(1);
    }

    if (!worksheet) {
      console.log("📄 Planilha vazia, será criada com os novos dados.");
      return [];
    }

    const result: Rows = [];
    let hasHeader = false;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      // Detecta automaticamente se a primeira linha é cabeçalho
      if (rowNumber === 1) {
        const firstCell = row.getCell(1).value?.toString().toLowerCase() || "";
        const secondCell = row.getCell(2).value?.toString().toLowerCase() || "";

        // Se contém "sku" e "drop", provavelmente é cabeçalho
        if (firstCell.includes("sku") && secondCell.includes("drop")) {
          hasHeader = true;
          console.log("📋 Cabeçalho detectado na primeira linha.");
          return;
        }
      }

      // Pula cabeçalho se foi detectado
      if (hasHeader && rowNumber === 1) return;

      const index = row.getCell(1).value?.toString()?.trim() || "";
      const skuVariation = row.getCell(2).value?.toString()?.trim() || "";
      const skuPai = row.getCell(2).value?.toString()?.trim() || "";
      const category = row.getCell(2).value?.toString()?.trim() || "";
      const drop = row.getCell(2).value?.toString()?.trim() || "";

      // Só adiciona se tiver SKU válido
      if (skuVariation) {
        result.push({ category, drop, index, skuPai, skuVariation });
      }
    });

    console.log(`📊 ${result.length} registros existentes encontrados.`);
    return result;
  } catch (error: any) {
    console.error("❌ Erro ao ler arquivo existente:", error.message);
    console.log("📄 Arquivo será recriado.");
    return [];
  }
}
