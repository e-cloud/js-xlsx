import { recordhopper } from './24_hoppers'
import { encode_cell } from './27_csfutils'

/* [MS-XLSB] 2.6.4.1 */
export function parse_BrtCalcChainItem$(data, length) {
    const out = {}
    out.i = data.read_shift(4)
    const cell = {}
    cell.r = data.read_shift(4)
    cell.c = data.read_shift(4)
    out.r = encode_cell(cell)
    const flags = data.read_shift(1)
    if (flags & 0x2) {
        out.l = '1'
    }
    if (flags & 0x8) {
        out.a = '1'
    }
    return out
}

/* 18.6 Calculation Chain */
export function parse_cc_bin(data, opts) {
    const out = []
    const pass = false
    recordhopper(data, function hopper_cc(val, R_n, RT) {
        switch (RT) {
            case 0x003F:
                /* 'BrtCalcChainItem$' */
                out.push(val)
                break

            default:
                if ((R_n || '').indexOf('Begin') > 0) {
                } else if ((R_n || '').indexOf('End') > 0) {
                } else if (!pass || opts.WTF) {
                    throw new Error(`Unexpected record ${RT} ${R_n}`)
                }
        }
    })
    return out
}

function write_cc_bin(data, opts) {
}
