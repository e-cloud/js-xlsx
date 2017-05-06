import { DENSE } from './03_consts'
import * as SSF from './10_ssf'
import { datenum } from './20_jsutils'

export function decode_row(rowstr: string): number {
    return parseInt(unfix_row(rowstr), 10) - 1
}

export function encode_row(row: number): string {
    return `${row + 1}`
}

export function fix_row(cstr: string): string {
    return cstr.replace(/([A-Z]|^)(\d+)$/, '$1$$$2')
}

export function unfix_row(cstr: string): string {
    return cstr.replace(/\$(\d+)$/, '$1')
}

export function decode_col(colstr: string): number {
    const c = unfix_col(colstr)
    let d = 0
    let i = 0
    for (; i !== c.length; ++i) {
        d = 26 * d + c.charCodeAt(i) - 64
    }
    return d - 1
}

export function encode_col(col: number): string {
    let s = ''
    for (++col; col; col = Math.floor((col - 1) / 26)) {
        s = String.fromCharCode((col - 1) % 26 + 65) + s
    }
    return s
}

export function fix_col(cstr: string): string {
    return cstr.replace(/^([A-Z])/, '$$$1')
}

export function unfix_col(cstr: string): string {
    return cstr.replace(/^\$([A-Z])/, '$1')
}

export function split_cell(cstr: string): Array<string> {
    return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, '$1,$2').split(',')
}

export function decode_cell(cstr: string): CellAddress {
    const splt = split_cell(cstr)
    return { c: decode_col(splt[0]), r: decode_row(splt[1]) }
}

export function encode_cell(cell: CellAddress): string {
    return encode_col(cell.c) + encode_row(cell.r)
}

export function fix_cell(cstr: string): string {
    return fix_col(fix_row(cstr))
}

export function unfix_cell(cstr: string): string {
    return unfix_col(unfix_row(cstr))
}

export function decode_range(range: string): Range {
    const x = range.split(':').map(decode_cell)
    return { s: x[0], e: x[x.length - 1] }
}

/*# if only one arg, it is assumed to be a Range.  If 2 args, both are cell addresses */
export function encode_range(cs: CellAddrSpec | Range, ce?: CellAddrSpec): string {
    if (typeof ce === 'undefined' || typeof ce === 'number') {
        /*:: if(!(cs instanceof Range)) throw "unreachable"; */
        return encode_range(cs.s, cs.e)
    }
    /*:: if((cs instanceof Range)) throw "unreachable"; */
    if (typeof cs !== 'string') {
        cs = encode_cell(cs)
    }
    if (typeof ce !== 'string') {
        ce = encode_cell(ce)
    }
    /*:: if(typeof cs !== 'string') throw "unreachable"; */
    /*:: if(typeof ce !== 'string') throw "unreachable"; */
    return cs == ce ? cs : `${cs}:${ce}`
}

export function safe_decode_range(range: string): Range {
    const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } }
    let idx = 0
    let i = 0
    let cc = 0
    const len = range.length
    for (idx = 0; i < len; ++i) {
        if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) {
            break
        }
        idx = 26 * idx + cc
    }
    o.s.c = --idx

    for (idx = 0; i < len; ++i) {
        if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) {
            break
        }
        idx = 10 * idx + cc
    }
    o.s.r = --idx

    if (i === len || range.charCodeAt(++i) === 58) {
        o.e.c = o.s.c
        o.e.r = o.s.r
        return o
    }

    for (idx = 0; i != len; ++i) {
        if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) {
            break
        }
        idx = 26 * idx + cc
    }
    o.e.c = --idx

    for (idx = 0; i != len; ++i) {
        if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) {
            break
        }
        idx = 10 * idx + cc
    }
    o.e.r = --idx
    return o
}

export function safe_format_cell(cell: Cell, v) {
    const q = cell.t == 'd' && v instanceof Date
    if (cell.z != null) {
        try {
            return cell.w = SSF.format(cell.z, q ? datenum(v) : v)
        } catch (e) {
            // todo: handle error
        }
    }
    try {
        return cell.w = SSF.format((cell.XF || {}).ifmt || (q ? 14 : 0), q ? datenum(v) : v)
    } catch (e) {
        return `${v}`
    }
}

export function format_cell(cell: Cell, v?, o?) {
    if (cell == null || cell.t == null || cell.t == 'z') {
        return ''
    }
    if (cell.w !== undefined) {
        return cell.w
    }
    if (cell.t == 'd' && !cell.z && o && o.dateNF) {
        cell.z = o.dateNF
    }
    if (v == undefined) {
        return safe_format_cell(cell, cell.v)
    }
    return safe_format_cell(cell, v)
}

export function sheet_to_workbook(sheet: Worksheet, opts): Workbook {
    const n = opts && opts.sheet ? opts.sheet : 'Sheet1'
    const sheets = {}
    sheets[n] = sheet
    return { SheetNames: [n], Sheets: sheets }
}

export function aoa_to_sheet(data: AOA, opts ?): Worksheet {
    const o = opts || {}
    if (DENSE != null && o.dense == null) {
        o.dense = DENSE
    }
    const ws: Worksheet = o.dense ? [] : {}

    const range: Range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } }

    for (let R = 0; R != data.length; ++R) {
        for (let C = 0; C != data[R].length; ++C) {
            if (typeof data[R][C] === 'undefined') {
                continue
            }
            const cell: Cell = { v: data[R][C] }

            if (range.s.r > R) {
                range.s.r = R
            }
            if (range.s.c > C) {
                range.s.c = C
            }
            if (range.e.r < R) {
                range.e.r = R
            }
            if (range.e.c < C) {
                range.e.c = C
            }
            if (cell.v === null) {
                if (!o.cellStubs) {
                    continue
                }
                cell.t = 'z'
            } else if (typeof cell.v === 'number') {
                cell.t = 'n'
            } else if (typeof cell.v === 'boolean') {
                cell.t = 'b'
            } else if (cell.v instanceof Date) {
                cell.z = o.dateNF || SSF._table[14]
                if (o.cellDates) {
                    cell.t = 'd'
                    cell.w = SSF.format(cell.z, datenum(cell.v))
                } else {
                    cell.t = 'n'
                    cell.v = datenum(cell.v)
                    cell.w = SSF.format(cell.z, cell.v)
                }
            } else {
                cell.t = 's'
            }
            if (o.dense) {
                if (!ws[R]) {
                    ws[R] = []
                }
                ws[R][C] = cell
            } else {
                const cell_ref = encode_cell({ c: C, r: R })
                ws[cell_ref] = cell
            }
        }
    }
    if (range.s.c < 10000000) {
        ws['!ref'] = encode_range(range)
    }
    return ws
}
