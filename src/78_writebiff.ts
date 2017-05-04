import { DENSE } from './03_consts'
import { datenum } from './20_jsutils'
/* BIFF2-4 single-sheet workbooks */
import { is_buf, new_buf } from './23_binutils'
import { buf_array } from './24_hoppers'
import { encode_col, encode_row, safe_decode_range } from './27_csfutils'

export function write_biff_rec(ba /*:BufArray*/, t /*:number*/, payload, length? /*:?number*/) {
    const len = length || (payload || []).length
    const o = ba.next(4 + len)
    o.write_shift(2, t)
    o.write_shift(2, len)
    if (/*:: len != null &&*/len > 0 && is_buf(payload)) {
        ba.push(payload)
    }
}

export function write_BOF(wb /*:Workbook*/, o) {
    if (o.bookType != 'biff2') {
        throw 'unsupported BIFF version'
    }
    const out = new_buf(4)
    out.write_shift(2, 0x0002) // "unused"
    out.write_shift(2, 0x0010) // Sheet
    return out
}

export function write_BIFF2Cell(out, r /*:number*/, c /*:number*/) {
    if (!out) out = new_buf(7)
    out.write_shift(2, r)
    out.write_shift(2, c)
    out.write_shift(1, 0)
    out.write_shift(1, 0)
    out.write_shift(1, 0)
    return out
}

export function write_BIFF2INT(r /*:number*/, c /*:number*/, val) {
    const out = new_buf(9)
    write_BIFF2Cell(out, r, c)
    out.write_shift(2, val)
    return out
}

export function write_BIFF2NUMBER(r, c, val) {
    const out = new_buf(15)
    write_BIFF2Cell(out, r, c)
    out.write_shift(8, val, 'f')
    return out
}

export function write_BIFF2BERR(r, c, val, t) {
    const out = new_buf(9)
    write_BIFF2Cell(out, r, c)
    if (t == 'e') {
        out.write_shift(1, val)
        out.write_shift(1, 1)
    } else {
        out.write_shift(1, val ? 1 : 0)
        out.write_shift(1, 0)
    }
    return out
}

/* TODO: codepage, large strings */
export function write_BIFF2LABEL(r, c, val) {
    const out = new_buf(8 + 2 * val.length)
    write_BIFF2Cell(out, r, c)
    out.write_shift(1, val.length)
    out.write_shift(val.length, val, 'sbcs')
    return out.l < out.length ? out.slice(0, out.l) : out
}

export function write_ws_biff_cell(ba /*:BufArray*/, cell /*:Cell*/, R /*:number*/, C /*:number*/, opts) {
    if (cell.v != null) {
        switch (cell.t) {
            case 'd':
            case 'n':
                const v = cell.t == 'd' ? datenum(cell.v) : cell.v
                if (v == (v | 0) && v >= 0 && v < 65536) {
                    write_biff_rec(ba, 0x0002, write_BIFF2INT(R, C, v))
                } else {
                    write_biff_rec(ba, 0x0003, write_BIFF2NUMBER(R, C, v))
                }
                return
            case 'b':
            case 'e':
                write_biff_rec(ba, 0x0005, write_BIFF2BERR(R, C, cell.v, cell.t))
                return
            /* TODO: codepage, sst */
            case 's':
            case 'str':
                write_biff_rec(ba, 0x0004, write_BIFF2LABEL(R, C, cell.v))
                return
        }
    }
    write_biff_rec(ba, 0x0001, write_BIFF2Cell(null, R, C))
}

export function write_biff_ws(ba /*:BufArray*/, ws /*:Worksheet*/, idx /*:number*/, opts, wb /*:Workbook*/) {
    const dense = Array.isArray(ws)
    const range = safe_decode_range(ws['!ref'] || 'A1')
    let ref
    let rr = ''
    const cols = []
    for (let R = range.s.r; R <= range.e.r; ++R) {
        rr = encode_row(R)
        for (let C = range.s.c; C <= range.e.c; ++C) {
            if (R === range.s.r) cols[C] = encode_col(C)
            ref = cols[C] + rr
            const cell = dense ? ws[R][C] : ws[ref]
            if (!cell) continue
            /* write cell */
            write_ws_biff_cell(ba, cell, R, C, opts)
        }
    }
}

/* Based on test files */
export function write_biff_buf(wb /*:Workbook*/, opts /*:WriteOpts*/) {
    const o = opts || {}
    if (DENSE != null && o.dense == null) o.dense = DENSE
    const ba = buf_array()
    let idx = 0
    for (let i = 0; i < wb.SheetNames.length; ++i) if (wb.SheetNames[i] == o.sheet) idx = i
    if (idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) {
        throw new Error(`Sheet not found: ${o.sheet}`)
    }
    write_biff_rec(ba, 0x0009, write_BOF(wb, o))
    /* ... */
    write_biff_ws(ba, wb.Sheets[wb.SheetNames[idx]], idx, o, wb)
    /* ... */
    write_biff_rec(ba, 0x000a)
    // TODO
    return ba.end()
}
