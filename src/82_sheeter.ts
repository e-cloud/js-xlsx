import { DIF, PRN, SYLK } from './40_harb'
/* actual implementation elsewhere, wrappers are for read/write */
import { HTML_ } from './79_html'
import { sheet_to_csv, sheet_to_txt } from './90_utils'

function write_obj_str(factory /*:WriteObjStrFactory*/) {
    return function write_str(wb /*:Workbook*/, o /*:WriteOpts*/) /*:string*/ {
        let idx = 0
        for (let i = 0; i < wb.SheetNames.length; ++i) {
            if (wb.SheetNames[i] == o.sheet) {
                idx = i
            }
        }
        if (idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) {
            throw new Error(`Sheet not found: ${o.sheet}`)
        }
        return factory.from_sheet(wb.Sheets[wb.SheetNames[idx]], o)
    }
}

export const write_htm_str = write_obj_str(HTML_)
export const write_csv_str = write_obj_str({ from_sheet: sheet_to_csv })
export const write_slk_str = write_obj_str(SYLK)
export const write_dif_str = write_obj_str(DIF)
export const write_prn_str = write_obj_str(PRN)
export const write_txt_str = write_obj_str({ from_sheet: sheet_to_txt })
