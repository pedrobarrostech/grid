import React, { useRef, useEffect, useCallback } from "react";
import CalcEngine, {
  CellConfig as CalcCellConfig,
  CellConfigGetter as CalcCellConfigGetter
} from "@rowsncolumns/calc";
import { CellInterface, castToString } from "@rowsncolumns/grid";
import {
  SheetID,
  CellsBySheet,
  CellConfig,
  FormulaMap,
  CellConfigBySheetNameGetter
} from "./../Spreadsheet";
import { formulas as defaultFormulas } from "../formulas";

export interface UseCalcOptions {
  formulas?: FormulaMap;
  getCellConfig: React.MutableRefObject<
    CellConfigBySheetNameGetter | undefined
  >;
}

const useCalc = ({ formulas, getCellConfig }: UseCalcOptions) => {
  const engine = useRef<CalcEngine>();
  useEffect(() => {
    engine.current = new CalcEngine({
      functions: {
        ...defaultFormulas,
        ...formulas
      }
    });
  }, []);

  const onCalculate = useCallback(
    (value: React.ReactText, sheet: SheetID, cell: CellInterface) => {
      const sheetId = castToString(sheet);
      if (!sheetId || !getCellConfig.current) return;
      return engine.current?.calculate(
        castToString(value) || "",
        sheetId,
        cell,
        getCellConfig.current as CalcCellConfigGetter
      );
    },
    []
  );

  const onCalculateBatch = useCallback(
    (changes: CellsBySheet, sheet: SheetID) => {
      const sheetId = castToString(sheet);
      if (!sheetId || !getCellConfig?.current) return;
      // @ts-ignore
      return engine.current?.calculateBatch(
        changes as Partial<CellConfig>,
        sheetId,
        getCellConfig.current as CalcCellConfigGetter
      );
    },
    []
  );

  const initializeEngine = useCallback((changes: CellsBySheet):
    | Promise<Partial<CellConfig> | undefined>
    | undefined => {
    if (!getCellConfig?.current) return;
    // @ts-ignore
    return engine.current?.initialize(
      changes as Partial<CellConfig>,
      getCellConfig.current as CalcCellConfigGetter
    );
  }, []);

  const getSupportedFormulas = useCallback(() => {
    return engine.current?.parser.formulaParser.supportedFunctions() ?? [];
  }, []);

  return {
    onCalculate,
    onCalculateBatch,
    initializeEngine,
    getSupportedFormulas
  };
};

export default useCalc;
