import * as fs from "fs";

// Função para verificar se o arquivo está em uso
export async function isFileInUse(filePath: string): Promise<boolean> {
  try {
    const fd = fs.openSync(filePath, "r+");
    fs.closeSync(fd);
    return false;
  } catch (error: any) {
    if (error.code === "EBUSY" || error.code === "EPERM") {
      return true;
    }
    return false;
  }
}
