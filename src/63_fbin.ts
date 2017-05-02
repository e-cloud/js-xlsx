/* XLSB Parsed Formula records have the same shape */
import { parse_RgbExtra, parse_Rgce } from './62_fxls'

function parse_XLSBParsedFormula(data, length, opts) {
    const end = data.l + length
    const cce = data.read_shift(4)
    const rgce = parse_Rgce(data, cce, opts)
    const cb = data.read_shift(4)
    const rgcb = cb > 0 ? parse_RgbExtra(data, cb, rgce, opts) : null
    return [rgce, rgcb]
}

/* [MS-XLSB] 2.5.97.1 ArrayParsedFormula */
export const parse_XLSBArrayParsedFormula = parse_XLSBParsedFormula
/* [MS-XLSB] 2.5.97.4 CellParsedFormula */
export const parse_XLSBCellParsedFormula = parse_XLSBParsedFormula
/* [MS-XLSB] 2.5.97.12 NameParsedFormula */
export const parse_XLSBNameParsedFormula = parse_XLSBParsedFormula
/* [MS-XLSB] 2.5.97.98 SharedParsedFormula */
export const parse_XLSBSharedParsedFormula = parse_XLSBParsedFormula
