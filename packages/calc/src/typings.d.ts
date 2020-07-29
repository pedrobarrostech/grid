declare module "fast-formula-parser" {
  class FormulaParser {
    constructor(options: any) {}
    parse: (text: string | null, position: any) => any;
    parseAsync: (text: string | null, position: any, array?: boolean) => Promise<any>;
    getValue: (
      sheet: Sheet,
      row: number,
      col: number
    ) => React.ReactText | undefined;
  }

  export default FormulaParser;
}

declare module "fast-formula-parser/grammar/dependency/hooks" {
  class DepParser {
    parse: (text: string, position: any) => any;
  }
}

declare module "fast-formula-parser/formulas/error" {
  class FormulaError {
    constructor(readonly error: string, readonly message?: string) {}
  }
  export default FormulaError;
}
