import FormulaError from "fast-formula-parser/formulas/error";

export interface FunctionArgument {
  value: string | number;
  isArray: boolean;
  isRangeRef: boolean;
  isCellRef: boolean;
}
export function importData(arg: FunctionArgument | undefined) {
  if (!arg || !arg.value)
    throw new FormulaError("#N/A", "Wrong number of arguments provided.");
  const { value } = arg;
  return fetch(value.toString())
    .then((r) => r.text())
    .then((response) => {
      const separator = value.toString().endsWith("tsv") ? "\t" : ",";
      const data = [];
      const rows = response.split("\n");
      for (const row of rows) {
        const cols = row.split(separator);
        data.push(cols);
      }
      return data;
    });
}

export function min(...arg: FunctionArgument[]) {
  return Math.min(...arg.map((item) => Number(item.value)));
}
export function max(...arg: FunctionArgument[]) {
  return Math.max(...arg.map((item) => Number(item.value)));
}

export function hyperLink(
  urlArg: FunctionArgument,
  titleArg: FunctionArgument
) {
  if ((titleArg?.value as any) instanceof FormulaError) {
    throw new FormulaError("#VALUE", "Invalid title");
  }
  if ((urlArg?.value as any) instanceof FormulaError) {
    throw new FormulaError("#VALUE", "Invalid url");
  }
  return JSON.stringify({
    title: titleArg?.value,
    hyperlink: urlArg?.value,
    datatype: "hyperlink",
  });
}

/* Default export */
export const formulas = {
  IMPORTDATA: importData,
  MIN: min,
  MAX: max,
  HYPERLINK: hyperLink,
};

export { FormulaError };
