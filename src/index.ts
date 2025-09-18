import { ensureFileAccess } from "./ensure-file-access";
import { readData } from "./read-data";
import { writeData } from "./write-data";
import { openExcelFile } from "./open-excel-file";

const filePath = "./SKU_DROP.xlsx";
const skus: Array<string> = [
  "TS751.01",
  "TS752.01",
  "TS753.01",
  "TS754.01",
  "TS755.01",
  "TS756.01",
  "TS757.01",
  "TS758.01",
  "TS759.01",
];

async function main(): Promise<void> {
  try {
    console.log("🚀 Iniciando processamento...");

    // Garante que o arquivo esteja acessível antes de escrever
    await ensureFileAccess(filePath);

    const currentRows = await readData({ filePath });
    const hasNewData = await writeData({ currentRows, filePath, skus });

    // Abre o arquivo Excel apenas se houver novos dados
    if (hasNewData) {
      await openExcelFile(filePath);
    } else {
      console.log(
        "ℹ️  Nenhum novo dado para processar. Planilha não foi alterada."
      );
    }
  } catch (error: any) {
    console.error("❌ Erro ao processar dados:", error.message);
    throw error;
  }
}

// Executa o script principal
main().catch(console.error);
