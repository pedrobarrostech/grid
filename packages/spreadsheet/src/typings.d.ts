declare module "fast-formula-parser/formulas/error" {
  class FormulaError {
    constructor(readonly error: string, readonly message?: string) {}
  }
  export default FormulaError;
}
