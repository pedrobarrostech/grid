import React, { useRef, useEffect, useCallback } from 'react'
import CalcEngine from '@rowsncolumns/calc'
import { CellInterface } from '@rowsncolumns/grid'
import { SheetID, CellConfigGetter, CellsBySheet }  from './../Spreadsheet'
import { castToString } from '../constants'

export interface UseCalcOptions {
  functions?: Record<string, (args: any) => void>;
  getCellConfig: React.MutableRefObject<CellConfigGetter | undefined>
}

const useCalc = ({ functions, getCellConfig }: UseCalcOptions)  => {
  const engine = useRef<CalcEngine>()
  useEffect(() => {
    engine.current = new CalcEngine({
      functions
    })
  }, [])

  const onCalculate = useCallback((value: React.ReactText, sheet: SheetID, cell: CellInterface) => {
    const sheetId = castToString(sheet)
    if (!sheetId || !getCellConfig.current) return
    return engine.current?.calculate(castToString(value), sheetId, cell, getCellConfig.current)
  }, [  ])

  const onCalculateBatch = useCallback((changes: CellsBySheet, sheet: SheetID) => {
    const sheetId = castToString(sheet)
    if (!sheetId || !getCellConfig?.current) return
    return engine.current?.calculateBatch(changes, sheetId, getCellConfig.current)
  }, [  ])

  const initializeEngine = useCallback((changes: CellsBySheet) => {
    if (!getCellConfig?.current) return
    return engine.current?.initialize(changes, getCellConfig.current)
  }, [])

  return {
    onCalculate,
    onCalculateBatch,
    initializeEngine
  }
}

export default useCalc