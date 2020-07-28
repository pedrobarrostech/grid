import {
  FormulaParser,
  Sheet,
  GetValue,
  CellConfig,
  CellRange
} from "./parser";
import { Dag, Node, DependencyMapping } from "./graph";
import { cellToAddress, isNull, createPosition } from "./helpers";
import merge from "lodash.merge";

interface CellInterface {
  rowIndex: number;
  columnIndex: number;
}
export type CellsBySheet = Record<string, Cells>;
type Cells = Record<string, Cell>;
type Cell = Record<string, CellConfig>;

/**
 * Todo
 * Remove dependencies on delete
 */
class CalcEngine {
  parser: FormulaParser;
  dag: Dag<Node>;
  mapping: DependencyMapping;
  constructor() {
    this.parser = new FormulaParser();
    this.dag = new Dag<Node>(node => node.children);
    this.mapping = new DependencyMapping();
  }

  /**
   * Parses a single cell
   * @param value
   * @param sheet
   * @param cell
   * @param getValue
   */
  calculate = async (
    value: string,
    sheet: Sheet,
    cell: CellInterface,
    getValue: GetValue
  ) => {
    const config = getValue(sheet, cell);
    if (config?.datatype !== "formula") {
      const cellAddress = cellToAddress(cell);
      if (!cellAddress) return;
      if (!this.mapping.has(cellAddress, sheet)) {
        console.log("No dependencies for ", cellAddress);
        return;
      }
      const node = this.mapping.get(cellAddress, sheet, cell);
      if (!node) return void 0;
      const dependencies = this.dag.visit([node]);

      if (dependencies.size > 0) {
        dependencies.delete(node);
        const values = await this.calculateDependencies(dependencies, getValue);
        return values;
      }
      return void 0;
    }

    const formula = value.substr(1);
    // @ts-ignore
    const cellAddress = cellToAddress(cell);
    // @ts-ignore
    const parentNode = this.mapping.get(cellAddress, sheet, cell);
    if (!cellAddress || !parentNode) {
      return void 0;
    }
    const position = createPosition(
      sheet,
      Number(cell.rowIndex),
      Number(cell.columnIndex)
    );
    const changes: CellsBySheet = {};
    changes[sheet] = changes[sheet] ?? {};
    changes[sheet][cell.rowIndex] = changes[sheet][cell.rowIndex] ?? {};

    /**
     * catch dependency errors
     */
    let dependencies;
    try {
      dependencies = this.parser.getDependencies(formula, position);
    } catch (err) {
      console.warn("Error parsing formula: ", formula, cell);
      changes[sheet][cell.rowIndex][cell.columnIndex] = {
        error: "Error parsing formula " + err.toString()
      };
      return changes;
    }

    /* Calculate */
    const result = await this.parser.parse(formula, position, getValue);

    /* Create results */
    changes[sheet] = changes[sheet] ?? {};
    changes[sheet][cell.rowIndex] = changes[sheet][cell.rowIndex] ?? {};
    changes[sheet][cell.rowIndex][cell.columnIndex] = result;

    /**
     * Cache current value
     * Eg
     * 1. User enters B1=SUM(A1, 20)
     * 2. B2 = B1
     * 3. When user changes B1, we need to use the result to calculate B2
     * 4. So we are merging the results to `getValue` in parser
     */
    this.parser.cacheValues(changes);

    /**
     * Add dependencies
     */
    for (const dep of dependencies) {
      if (dep.from) {
        const { from, to, sheet } = dep as CellRange;
        for (let i = from.row; i <= to.row; i++) {
          for (let j = from.col; j <= to.col; j++) {
            const cell = { rowIndex: i, columnIndex: j };
            const address = cellToAddress(cell);
            if (!address) {
              continue;
            }
            const node = this.mapping.get(address, sheet, cell);
            node?.children.add(parentNode);
          }
        }
      } else {
        const { row, col, sheet, address } = dep;
        const cellNode = { rowIndex: row, columnIndex: col };
        const node = this.mapping.get(address, sheet, cellNode);
        node?.children.add(parentNode);
      }
    }

    /**
     * Visit dependents
     */
    const directDependencies = this.dag.visit([parentNode]);
    let values = {};
    if (directDependencies.size > 0) {
      directDependencies.delete(parentNode);
      values = await this.calculateDependencies(directDependencies, getValue);
    }

    /* Remove all caches after calculation is complete */
    this.parser.clearCachedValues();

    return merge(changes, values);
  };

  calculateDependencies = async (
    dependencies: Set<Node>,
    getValue: GetValue
  ) => {
    const values: CellsBySheet = {};
    for (const { cell, address, sheet } of dependencies) {
      const config = getValue(sheet, cell);
      const isFormula = config?.datatype === "formula";
      if (!isFormula || isNull(config.text) || config.text === void 0) continue;
      const position = createPosition(
        sheet,
        Number(cell.rowIndex),
        Number(cell.columnIndex)
      );
      const formula = config.text.substr(1) ?? null;
      const result = await this.parser.parse(formula, position, getValue);
      values[sheet] = values[sheet] ?? {};
      values[sheet][cell.rowIndex] = values[sheet][cell.rowIndex] ?? {};
      values[sheet][cell.rowIndex][cell.columnIndex] = result;

      /* Cache current values so subsequent calculations can use the result */
      this.parser.cacheValues(values);
    }
    return values;
  };

  calculateBatch = async (
    changes: CellsBySheet,
    sheet: Sheet,
    getValue: GetValue
  ) => {};

  /**
   * Set dependencies in graph
   * @param changes
   */
  initialize = async (changes: CellsBySheet, getValue: GetValue) => {
    const values = {};
    for (const sheet in changes) {
      for (const rowIndex in changes[sheet]) {
        for (const columnIndex in changes[sheet][rowIndex]) {
          const cellConfig = changes[sheet][rowIndex][columnIndex];
          const datatype = cellConfig.datatype;
          const text = cellConfig?.text as string;

          if (!isNull(text) && datatype === "formula") {
            const formula = text.substr(1);
            const cell = {
              rowIndex: Number(rowIndex),
              columnIndex: Number(columnIndex)
            };
            const cellAddress = cellToAddress(cell);
            if (!cellAddress) {
              continue;
            }

            const parentNode = this.mapping.get(cellAddress, sheet, cell);
            if (!cellAddress || !parentNode) {
              continue;
            }
            const position = createPosition(
              sheet,
              Number(rowIndex),
              Number(columnIndex)
            );

            const result = await this.calculate(text, sheet, cell, getValue);
            /* Merge results  */
            merge(values, result);

            try {
              const dependents = this.parser.getDependencies(formula, position);
              for (const dep of dependents) {
                if (dep.from) {
                  const { from, to, sheet } = dep as CellRange;
                  for (let i = from.row; i <= to.row; i++) {
                    for (let j = from.col; j <= to.col; j++) {
                      const cell = { rowIndex: i, columnIndex: j };
                      const address = cellToAddress(cell);
                      if (!address) {
                        continue;
                      }
                      const node = this.mapping.get(address, sheet, cell);
                      node?.children.add(parentNode);
                    }
                  }
                } else {
                  const { row, col, sheet, address } = dep;
                  const cell = { rowIndex: row, columnIndex: col };
                  const node = this.mapping.get(address, sheet, cell);
                  node?.children.add(parentNode);
                }
              }
            } catch (err) {
              console.log("Error parsing formula ", err);
            }
          }
        }
      }

      return values;
    }
  };
}

export default CalcEngine;
