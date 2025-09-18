import { Rows } from "./types";
import { workbook } from "./workbook";

interface Input {
  filePath: string;
  currentRows: Rows;
  skus: Array<string>;
}

export async function writeData(input: Input): Promise<void> {
  try {
    const newSkus = input.skus.filter(
      (sku) => !input.currentRows.some((item) => item.sku === sku)
    );

    const rows: Rows = [...input.currentRows];

    if (newSkus.length === 0) {
      console.log("ℹ️  Nenhum SKU novo para buscar.");
      return;
    }

    for (const sku of newSkus) {
      console.log(`🔍 Buscando: ${sku}`);
      const url = `https://app.asana.com/api/1.0/workspaces/1202771719354212/tasks/search?opt_fields=custom_fields,name&custom_fields.1204602756616122.value=${sku}`;

      const response = await fetch(url, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization:
            "Bearer 2/1202771717967620/1206382577186228:7efb45b76d8d3b0cfa12529c5bd9a9d2",
        },
      });

      const json = await response.json();
      if (!json || !json.data || json.data?.length === 0) continue;

      json.data.forEach((product: any) => {
        const drop = product.custom_fields.find(
          (item: any) => item.name.toUpperCase() === "DROP"
        );
        const grid = product.custom_fields.find(
          (item: any) => item.name.toUpperCase() === "GRADE"
        )?.display_value;

        if (grid) {
          grid.split("-").forEach((size: string) => {
            const row = {
              sku: `${sku}-${size}`,
              drop: drop?.display_value,
            };
            console.log(`  ➕ Adicionando: ${row.sku} - DROP: ${row.drop}`);
            rows.push(row);
          });
        }
      });
    }

    console.log("💾 Salvando dados no Excel...");

    workbook.views = [
      {
        x: 0,
        y: 0,
        width: 10000,
        height: 20000,
        firstSheet: 0,
        activeTab: 1,
        visibility: "visible",
      },
    ];

    const sheetName = "SKU x DROP";
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      { header: "SKU", key: "sku", width: 30 },
      { header: "DROP", key: "drop", width: 30 },
    ];

    rows.forEach((item) => sheet.addRow({ sku: item.sku, drop: item.drop }));

    await workbook.xlsx.writeFile(input.filePath);
    console.log("✅ Arquivo Excel atualizado com sucesso!");
    console.log(`📊 Total de registros: ${rows.length}`);
  } catch (error: any) {
    console.error("❌ Erro ao processar dados:", error.message);
    throw error;
  }
}
