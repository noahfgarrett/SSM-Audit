import test from 'node:test'
import assert from 'node:assert/strict'

import { sheetAoaAsync } from '../src/io/workbook.js'

test('workbook reader supports modern SheetJS dense worksheet storage', async () => {
  globalThis.XLSX={utils:{format_cell:cell=>String(cell.v??'')}}
  const sheet={'!ref':'A1:B2','!data':[
    [{v:'Equipment ID',w:'Equipment ID'},{v:'UPN',w:'UPN'}],
    [{v:'B1-PUMP-01',w:'B1-PUMP-01'},{v:101,w:'101'}],
  ]}
  const parsed=await sheetAoaAsync(sheet)
  assert.deepEqual(parsed.aoa,[['Equipment ID','UPN'],['B1-PUMP-01','101']])
  assert.deepEqual(parsed.rowNums,[0,1])
})
