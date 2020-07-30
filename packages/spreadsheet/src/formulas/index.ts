import FormulaError from "fast-formula-parser/formulas/error";

export interface FunctionArgument {
  value: string;
  isArray: boolean;
  isRangeRef: boolean;
  isCellRef: boolean;
}
export function importData(arg: FunctionArgument | undefined) {
  if (!arg)
    throw new FormulaError("#N/A", "Wrong number of arguments provided.");
  const { value } = arg;
  return fetch(value)
    .then((r) => r.text())
    .then((response) => {
      const data = [];
      const rows = response.split("\n");
      for (const row of rows) {
        const cols = row.split(",");
        data.push(cols);
      }
      return data;
    });
}

/* Default export */
export const formulas = {
  IMPORTDATA: importData,
};

export { FormulaError };
