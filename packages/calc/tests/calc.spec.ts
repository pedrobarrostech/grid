// @ts-nocheck
import CalcEngine, { CellsBySheet } from './../src/calc'
import { Node } from './../src/graph'
import { CellInterface } from '../src/parser'

describe('Calc Engine', () => {
  let engine: CalcEngine
  beforeEach(() => {
    engine = new CalcEngine()
  })
  it('exists', () => {
    expect(CalcEngine).toBeDefined()
  })

  it('can initialize', async () => {
    const data: CellsBySheet = {
      'sheet 1': {
        1: {
          2: {
            text: '=SUM(A1, A2)',
            datatype: 'formula'
          }
        }
      }
    }
    await engine.initialize(data, () => {
      return {
        text: '12'
      }
    })
    const nodeB1 = engine.mapping.get('B1', 'sheet 1', { rowIndex: 1, columnIndex: 2}) as Node
    expect(engine.mapping.get('A1', 'sheet 1', { rowIndex: 1, columnIndex: 1})?.children.has(nodeB1)).toBe(true)
  })

  it('can calculate', async () => {
    const sheet = 'sheet 1'
    const cell = { rowIndex: 1, columnIndex: 2}
    const getValue = (sheet: string, cell: CellInterface) => {
      if (cell.rowIndex === 1 && cell.columnIndex === 2) {
        return {
          text: '=SUM(A1, 20)',
          datatype: 'formula',
        }
      }
      return {
        text: '200',
        datatype: 'number'
      }
    }    
    const results = await engine.calculate('=SUM(A1, 20)', sheet, cell, getValue)
    expect(results[sheet][cell.rowIndex][cell.columnIndex].result).toBe(220)
  })

  it('can do nested calculations', async () => {
    const sheet = 'sheet 1'
    const cell = { rowIndex: 1, columnIndex: 2}
    const getValue = (sheet: string, cell: CellInterface) => {
      if (cell.rowIndex === 1 && cell.columnIndex === 2) {
        return {
          text: '=SUM(A1, B2)',
          datatype: 'formula',
        }
      }
      if (cell.rowIndex === 2 && cell.columnIndex === 2) {
        return {
          text: '=SUM(C2, 10)',
          datatype: 'formula',
          result: 20
        }
      }
      return {
        text: '200',
        datatype: 'number'
      }
    }    
    const results = await engine.calculate('=SUM(A1, B2)', sheet, cell, getValue)
    expect(results[sheet][cell.rowIndex][cell.columnIndex].result).toBe(220)
  })
})