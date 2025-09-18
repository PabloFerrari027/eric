import * as fs from "fs";
import { Rows } from "./types";
import { workbook } from "./workbook";

interface Input {
  filePath: string;
}

export async function readData(input: Input): Promise<Rows> {
  const isFileExists = fs.existsSync(input.filePath);

  if (!isFileExists) return [];

  await workbook.xlsx.readFile(input.filePath);

  const worksheet = workbook.getWorksheet(1);

  if (!worksheet) return [];

  const result: Rows = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const sku = row.getCell(1).value?.toString() || "";
    const drop = row.getCell(2).value?.toString() || undefined;
    result.push({ sku, drop });
  });

  return result;
}
