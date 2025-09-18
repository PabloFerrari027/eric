import { ensureFileAccess } from "./ensure-file-access";
import { readData } from "./read-data";
import { writeData } from "./write-data";

const filePath = "./SKU_DROP.xlsx";
const skus: Array<string> = ["TS756.01", "TS755.01"];

async function main(): Promise<void> {
  try {
    console.log("🚀 Iniciando processamento...");
    // Garante que o arquivo esteja acessível antes de escrever
    await ensureFileAccess(filePath);
    const currentRows = await readData({ filePath });
    await writeData({ currentRows, filePath, skus });
  } catch (error: any) {
    console.error("❌ Erro ao processar dados:", error.message);
    throw error;
  }
}

// Executa o script principal
main().catch(console.error);
