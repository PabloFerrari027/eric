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

    console.log("💾 Atualizando planilha preservando formatações...");

    // Carrega a planilha existente OU cria uma nova
    const workbook = new Workbook();
    let worksheet;

    if (fs.existsSync(input.filePath)) {
      // Carrega a planilha existente mantendo TODAS as formatações
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
        { header: "SKU", key: "sku", width: 30 },
        { header: "DROP", key: "drop", width: 30 },
      ];

      // Formata cabeçalho
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE6E6FA" },
      };
    }

    // Encontra a próxima linha vazia (após os dados existentes)
    let lastRow = 1; // Começa após o cabeçalho
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > lastRow) lastRow = rowNumber;
    });

    // Adiciona apenas as novas linhas, preservando formatações existentes
    newRows.forEach((newRow, index) => {
      const rowNumber = lastRow + 1 + index;
      const row = worksheet.getRow(rowNumber);

      // Define apenas os valores, preserva qualquer formatação existente
      row.getCell(1).value = newRow.sku;
      row.getCell(2).value = newRow.drop;

      // Se for uma nova linha, aplica formatação básica apenas nas células com dados
      if (rowNumber > worksheet.actualRowCount) {
        row.getCell(1).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        row.getCell(2).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }

      row.commit();
    });

    // Salva mantendo TODAS as formatações existentes
    await workbook.xlsx.writeFile(input.filePath);

    console.log("✅ Planilha atualizada preservando formatações existentes!");
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
