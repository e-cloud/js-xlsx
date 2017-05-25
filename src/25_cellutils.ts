/* XLS ranges enforced */
import { dup } from './20_jsutils'
import { encode_cell, encode_col, encode_row, fix_col, fix_row } from './27_csfutils'

export function shift_cell_xls(cell: CellAddress, tgt, opts?): CellAddress {
    const out = dup(cell)
    if (tgt.s) {
        if (out.cRel) {
            out.c += tgt.s.c
        }
        if (out.rRel) {
            out.r += tgt.s.r
        }
    } else {
        out.c += tgt.c
        out.r += tgt.r
    }
    if (!opts || opts.biff < 12) {
        while (out.c >= 0x100) {
            out.c -= 0x100
        }
        while (out.r >= 0x10000) {
            out.r -= 0x10000
        }
    }
    return out
}

export function shift_range_xls(cell, range, opts) {
    const out = dup(cell)
    out.s = shift_cell_xls(out.s, range.s, opts)
    out.e = shift_cell_xls(out.e, range.s, opts)
    return out
}

export function encode_cell_xls(c: CellAddress): string {
    let s = encode_cell(c)
    if (c.cRel === 0) {
        s = fix_col(s)
    }
    if (c.rRel === 0) {
        s = fix_row(s)
    }
    return s
}

export function encode_range_xls(r, opts): string {
    if (r.s.r == 0 && !r.s.rRel) {
        if (r.e.r == opts.biff >= 12 ? 0xFFFFF : 0xFFFF && !r.e.rRel) {
            return `${(r.s.cRel ? '' : '$') + encode_col(r.s.c)}:${r.e.cRel ? '' : '$'}${encode_col(r.e.c)}`
        }
    }
    if (r.s.c == 0 && !r.s.cRel) {
        if (r.e.c == opts.biff >= 12 ? 0xFFFF : 0xFF && !r.e.cRel) {
            return `${(r.s.rRel ? '' : '$') + encode_row(r.s.r)}:${r.e.rRel ? '' : '$'}${encode_row(r.e.r)}`
        }
    }
    return `${encode_cell_xls(r.s)}:${encode_cell_xls(r.e)}`
}
