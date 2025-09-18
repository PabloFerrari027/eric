import * as fs from "fs";
import * as path from "path";
import { isFileInUse } from "./is-file-in-use";
import { closeSpecificExcelFile } from "./close-specific-excel-file";

// Função principal para garantir acesso ao arquivo
export async function ensureFileAccess(
  filePath: string,
  maxAttempts: number = 3
): Promise<void> {
  const fileName = path.basename(filePath);

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      if (!fs.existsSync(filePath)) {
        console.log("Arquivo não existe ainda, será criado.");
        return;
      }

      const isInUse = await isFileInUse(filePath);

      if (!isInUse) {
        console.log(`✅ Arquivo ${fileName} disponível para uso.`);
        return;
      }

      console.log(
        `⚠️  Tentativa ${attempt}/${maxAttempts}: Arquivo ${fileName} está em uso...`
      );

      // Tenta fechar o arquivo específico
      console.log(`🔄 Tentando fechar ${fileName} no Excel...`);
      let closed = await closeSpecificExcelFile(fileName);

      if (closed) {
        console.log(`✅ Arquivo ${fileName} fechado com sucesso!`);
      } else {
        console.log(`❌ Não foi possível fechar ${fileName} automaticamente.`);
      }

      // Aguarda antes da próxima verificação
      await new Promise((resolve) => setTimeout(resolve, 2000));
    } catch (error: any) {
      console.log(`Erro na tentativa ${attempt}:`, error.message);
    }
  }

  // Última verificação
  if (await isFileInUse(filePath)) {
    console.log("\n❌ ATENÇÃO: O arquivo ainda está em uso!");
    console.log(
      `Por favor, feche manualmente o arquivo ${fileName} no Excel e execute o script novamente.`
    );
    throw new Error(
      `Arquivo ${fileName} ainda está sendo usado após ${maxAttempts} tentativas.`
    );
  }
}
