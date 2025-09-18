import { exec } from "child_process";
import { promisify } from "util";

const execAsync = promisify(exec);

// Função para fechar apenas a janela específica do Excel
export async function closeSpecificExcelFile(
  fileName: string
): Promise<boolean> {
  try {
    if (process.platform === "win32") {
      // PowerShell script para fechar apenas a janela específica
      const powershellScript = `
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel
        $excel = $null
        try {
          $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
          foreach ($workbook in $excel.Workbooks) {
            if ($workbook.Name -eq "${fileName}") {
              Write-Host "Fechando arquivo: $($workbook.Name)"
              $workbook.Close($false)
              return $true
            }
          }
          Write-Host "Arquivo ${fileName} não encontrado nas janelas abertas do Excel"
          return $false
        }
        catch {
          Write-Host "Excel não está executando ou erro: $($_.Exception.Message)"
          return $false
        }
        finally {
          if ($excel -ne $null) {
            [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
          }
        }
      `;

      const result = await execAsync(
        `powershell -Command "${powershellScript.replace(/\n/g, "; ")}"`
      );
      console.log("PowerShell output:", result.stdout);

      // Aguarda um momento para garantir que o arquivo foi fechado
      await new Promise((resolve) => setTimeout(resolve, 1500));
      return true;
    } else if (process.platform === "darwin") {
      // macOS - usando AppleScript
      const appleScript = `
        tell application "Microsoft Excel"
          set workbookList to every workbook
          repeat with wb in workbookList
            if name of wb is "${fileName}" then
              close wb saving no
              return true
            end if
          end repeat
          return false
        end tell
      `;

      await execAsync(`osascript -e '${appleScript}'`);
      console.log(`Tentativa de fechar ${fileName} no macOS concluída.`);
      await new Promise((resolve) => setTimeout(resolve, 1500));
      return true;
    } else {
      console.log(
        "Sistema operacional não suportado para fechamento automático do Excel."
      );
      return false;
    }
  } catch (error: any) {
    console.log(
      "Erro ao tentar fechar arquivo específico do Excel:",
      error.message
    );
    return false;
  }
}
