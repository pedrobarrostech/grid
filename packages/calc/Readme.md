## Calc

A calculation module for spreadsheet

### Targeted features

1. Runs asynchronously
1. Able to run on web workers
1. Able to return multiple results based on cell-dependency
1. Able to run multiple calculations (batch)
1. Able to run single calculation


```
import { calculate } from '@rowsncolumns/spreadsheet'

// Single cell calculation
const results = await calcEngine.calculate(value, sheet, cell, getCellConfig)

// Multiple batch
const results = await calcEngine.calculateBatch(changes, getCellConfig)
```