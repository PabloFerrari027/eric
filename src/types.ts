export interface Row {
  index: string
  skuVariation: string;
  skuPai: string;
  category: string
  drop: string | undefined;
}

export type Rows = Array<Row>;
