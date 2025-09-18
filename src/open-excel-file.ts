import { exec } from "child_process";
import { promisify } from "util";
import * as path from "path";

const execAsync = promisify(exec);

export async function openExcelFile(filePath: string): Promise<void> {
  try {
    const absolutePath = path.resolve(filePath);

    console.log(`📂 Abrindo arquivo: ${path.basename(filePath)}`);

    if (process.platform === "win32") {
      // Windows - abre o arquivo com o aplicativo padrão (Excel)
      await execAsync(`start "" "${absolutePath}"`);
    } else if (process.platform === "darwin") {
      // macOS - abre o arquivo com o aplicativo padrão
      await execAsync(`open "${absolutePath}"`);
    } else if (process.platform === "linux") {
      // Linux - abre o arquivo com o aplicativo padrão
      await execAsync(`xdg-open "${absolutePath}"`);
    } else {
      console.log(
        "❌ Sistema operacional não suportado para abertura automática."
      );
      return;
    }

    console.log("✅ Arquivo Excel aberto com sucesso!");
  } catch (error: any) {
    console.error("❌ Erro ao abrir o arquivo Excel:", error.message);
    console.log("💡 Você pode abrir o arquivo manualmente.");
  }
}
