import { Rows } from "./types";
import { Workbook } from "exceljs";
import * as fs from "fs";
import { FetchData } from "./fetch-data";

interface Input {
  filePath: string;
  currentRows: Rows;
  skus: Array<string>;
}

export async function writeData(input: Input): Promise<boolean> {
  try {
    // Busca os dados dos novos SKUs
    const newRows = await FetchData({
      currentRows: input.currentRows,
      filePath: input.filePath,
      skus: input.skus,
    });

    if (newRows.length === 0) {
      console.log("ℹ️  Nenhum dado novo encontrado na API.");
      return false;
    }

    console.log("💾 Atualizando planilha...");

    // Carrega a planilha existente OU cria uma nova
    const workbook = new Workbook();
    let worksheet;

    if (fs.existsSync(input.filePath)) {
      // Carrega a planilha existente
      await workbook.xlsx.readFile(input.filePath);
      worksheet =
        workbook.getWorksheet("SKU x DROP") || workbook.getWorksheet(1);

      if (!worksheet) {
        worksheet = workbook.addWorksheet("SKU x DROP");
        // Adiciona cabeçalhos apenas se for uma nova aba
        worksheet.getRow(1).values = ["SKU", "DROP"];
      }
    } else {
      // Cria nova planilha
      worksheet = workbook.addWorksheet("SKU x DROP");

      // Configura cabeçalhos com formatação básica
      worksheet.columns = [
        { header: "Index", key: "index", width: 30 },
        { header: "Código (SKU)", key: "skuVariation", width: 30 },
        { header: "sku s tam", key: "skuPai", width: 30 },
        { header: "Categoria", key: "category", width: 30 },
        { header: "DROP", key: "drop", width: 30 },
      ];

    }

    // Encontra a próxima linha vazia (após os dados existentes)
    let lastRow = 1; // Começa após o cabeçalho
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > lastRow) lastRow = rowNumber;
    });

    // Adiciona apenas as novas linhas
    newRows.forEach((newRow, index) => {
      const rowNumber = lastRow + 1 + index;
      const row = worksheet.getRow(rowNumber);

      // Define apenas os valores
      row.getCell(1).value = newRow.index;
      row.getCell(2).value = newRow.skuVariation;
      row.getCell(3).value = newRow.skuPai;
      row.getCell(4).value = newRow.category;
      row.getCell(5).value = newRow.drop;
      row.commit();
    });

    // Salva
    await workbook.xlsx.writeFile(input.filePath);

    console.log("✅ Planilha atualizada!");
    console.log(`📊 Novos registros adicionados: ${newRows.length}`);
    console.log(
      `📋 Total de registros: ${input.currentRows.length + newRows.length}`
    );

    return true;
  } catch (error: any) {
    console.error("❌ Erro ao processar dados:", error.message);
    throw error;
  }
}
