import * as cptable from 'codepage/dist/cpexcel.full.js'
import * as SSF from './10_ssf'
import { datenum } from './20_jsutils'
import {
    aoa_to_sheet,
    decode_cell,
    decode_col,
    decode_range,
    decode_row,
    encode_cell,
    encode_col,
    encode_range,
    encode_row,
    format_cell,
    safe_decode_range,
    split_cell,
} from './27_csfutils'
import { HTML_, parse_dom_table, table_to_book } from './79_html'

function sheet_to_json(sheet: Worksheet, opts ?: Sheet2JSONOpts) {
    if (sheet == null || sheet['!ref'] == null) {
        return []
    }
    let val = { t: 'n', v: 0 }
    let header = 0
    let offset = 1
    const hdr: Array<any> = []
    let isempty = true
    let v = 0
    let vv = ''
    let r = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
    const o = opts != null ? opts : {}
    const raw = o.raw
    const defval = o.defval
    const range = o.range != null ? o.range : sheet['!ref']
    if (o.header === 1) {
        header = 1
    } else if (o.header === 'A') {
        header = 2
    } else if (Array.isArray(o.header)) {
        header = 3
    }
    switch (typeof range) {
        case 'string':
            r = safe_decode_range(range)
            break
        case 'number':
            r = safe_decode_range(sheet['!ref'])
            r.s.r = range
            break
        default:
            r = range
    }
    if (header > 0) {
        offset = 0
    }
    let rr = encode_row(r.s.r)
    const cols = new Array(r.e.c - r.s.c + 1)
    const out = new Array(r.e.r - r.s.r - offset + 1)
    let outi = 0
    let counter = 0
    const dense = Array.isArray(sheet)
    let R = r.s.r
    let C = 0
    let CC = 0
    if (dense && !sheet[R]) {
        sheet[R] = []
    }
    for (C = r.s.c; C <= r.e.c; ++C) {
        cols[C] = encode_col(C)
        val = dense ? sheet[R][C] : sheet[cols[C] + rr]
        switch (header) {
            case 1:
                hdr[C] = C - r.s.c
                break
            case 2:
                hdr[C] = cols[C]
                break
            case 3:
                hdr[C] = o.header[C - r.s.c]
                break
            default:
                if (val == null) {
                    continue
                }
                vv = v = format_cell(val, null, o)
                counter = 0
                for (CC = 0; CC < hdr.length; ++CC) {
                    if (hdr[CC] == vv) {
                        vv = `${v}_${++counter}`
                    }
                }
                hdr[C] = vv
        }
    }
    let row = header === 1 ? [] : {}
    for (R = r.s.r + offset; R <= r.e.r; ++R) {
        rr = encode_row(R)
        isempty = true
        if (header === 1) {
            row = []
        } else {
            row = {}
            if (Object.defineProperty) {
                try {
                    Object.defineProperty(row, '__rowNum__', { value: R, enumerable: false })
                } catch (e) {
                    row.__rowNum__ = R
                }
            } else {
                row.__rowNum__ = R
            }
        }
        if (!dense || sheet[R]) {
            for (C = r.s.c; C <= r.e.c; ++C) {
                val = dense ? sheet[R][C] : sheet[cols[C] + rr]
                if (val === undefined || val.t === undefined) {
                    if (defval === undefined) {
                        continue
                    }
                    if (hdr[C] != null) {
                        row[hdr[C]] = defval
                        isempty = false
                    }
                    continue
                }
                v = val.v
                switch (val.t) {
                    case 'z':
                        if (v == null) {
                            break
                        }
                        continue
                    case 'e':
                        continue
                    case 's':
                    case 'd':
                    case 'b':
                    case 'n':
                        break
                    default:
                        throw new Error(`unrecognized type ${val.t}`)
                }
                if (hdr[C] != null) {
                    if (v == null) {
                        if (defval !== undefined) {
                            row[hdr[C]] = defval
                        } else if (raw && v === null) {
                            row[hdr[C]] = null
                        } else {
                            continue
                        }
                    } else {
                        row[hdr[C]] = raw ? v : format_cell(val, v, o)
                    }
                    isempty = false
                }
            }
        }
        if (isempty === false || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
            out[outi++] = row
        }
    }
    out.length = outi
    return out
}

const qreg = /"/g
export function make_csv_row(
    sheet: Worksheet,
    r: Range,
    R: number,
    cols: Array<string>,
    fs: number,
    rs: number,
    FS: string,
    o: Sheet2CSVOpts,
): string {
    let isempty = true
    let row = ''
    let txt = ''
    const rr = encode_row(R)
    for (let C = r.s.c; C <= r.e.c; ++C) {
        const val = o.dense ? (sheet[R] || [])[C] : sheet[cols[C] + rr]
        if (val == null) {
            txt = ''
        } else if (val.v != null) {
            isempty = false
            txt = `${format_cell(val, null, o)}`
            for (let i = 0, cc = 0; i !== txt.length; ++i) {
                if ((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {
                    txt = `"${txt.replace(qreg, '""')}"`
                    break
                }
            }
            if (txt == 'ID') {
                txt = '"ID"'
            }
        } else if (val.f != null && !val.F) {
            isempty = false
            txt = `=${val.f}`
            if (txt.includes(',')) {
                txt = `"${txt.replace(qreg, '""')}"`
            }
        } else {
            txt = ''
        }
        /* NOTE: Excel CSV does not support array formulae */
        row += (C === r.s.c ? '' : FS) + txt
    }
    if (o.blankrows === false && isempty) {
        return null
    }
    return row
}

export function sheet_to_csv(sheet: Worksheet, opts ?: Sheet2CSVOpts): string {
    const out = []
    const o = opts == null ? {} : opts
    if (sheet == null || sheet['!ref'] == null) {
        return ''
    }
    const r = safe_decode_range(sheet['!ref'])
    const FS = o.FS !== undefined ? o.FS : ','
    const fs = FS.charCodeAt(0)
    const RS = o.RS !== undefined ? o.RS : '\n'
    const rs = RS.charCodeAt(0)
    const endregex = new RegExp(`${FS == '|' ? '\\|' : FS}+$`)
    let row = ''
    const cols = []
    o.dense = Array.isArray(sheet)
    for (let C = r.s.c; C <= r.e.c; ++C) {
        cols[C] = encode_col(C)
    }
    for (let R = r.s.r; R <= r.e.r; ++R) {
        row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o)
        if (row == null) {
            continue
        }
        if (o.strip) {
            row = row.replace(endregex, '')
        }
        out.push(row + RS)
    }
    delete o.dense
    return out.join('')
}

export function sheet_to_txt(sheet: Worksheet, opts ?: Sheet2CSVOpts) {
    if (!opts) {
        opts = {}
    }
    opts.FS = '\t'
    opts.RS = '\n'
    const s = sheet_to_csv(sheet, opts)
    if (typeof cptable == 'undefined') {
        return s
    }
    const o = cptable.utils.encode(1200, s)
    return `\xFF\xFE${o}`
}

function sheet_to_formulae(sheet: Worksheet): Array<string> {
    let y = ''
    let x
    let val = ''
    if (sheet == null || sheet['!ref'] == null) {
        return []
    }
    const r = safe_decode_range(sheet['!ref'])
    let rr = ''
    const cols = []
    let C
    const cmds = new Array((r.e.r - r.s.r + 1) * (r.e.c - r.s.c + 1))
    let i = 0
    const dense = Array.isArray(sheet)
    for (C = r.s.c; C <= r.e.c; ++C) {
        cols[C] = encode_col(C)
    }
    for (let R = r.s.r; R <= r.e.r; ++R) {
        rr = encode_row(R)
        for (C = r.s.c; C <= r.e.c; ++C) {
            y = cols[C] + rr
            x = dense ? (sheet[R] || [])[C] : sheet[y]
            val = ''
            if (x === undefined) {
                continue
            } else if (x.F != null) {
                y = x.F
                if (!x.f) {
                    continue
                }
                val = x.f
                if (!y.includes(':')) {
                    y = `${y}:${y}`
                }
            }
            if (x.f != null) {
                val = x.f
            } else if (x.t == 'z') {
                continue
            } else if (x.t == 'n' && x.v != null) {
                val = `${x.v}`
            } else if (x.t == 'b') {
                val = x.v ? 'TRUE' : 'FALSE'
            } else if (x.w !== undefined) {
                val = `'${x.w}`
            } else if (x.v === undefined) {
                continue
            } else if (x.t == 's') {
                val = `'${x.v}`
            } else {
                val = `${x.v}`
            }
            cmds[i++] = `${y}=${val}`
        }
    }
    cmds.length = i
    return cmds
}

function json_to_sheet(js: Array<any>, opts): Worksheet {
    const o = opts || {}
    const ws = {}
    let cell: Cell
    const range: Range = { s: { c: 0, r: 0 }, e: { c: 0, r: js.length } }
    const hdr = o.header || []
    let C = 0

    for (let R = 0; R != js.length; ++R) {
        Object.keys(js[R])
            .filter(x => js[R].hasOwnProperty(x))
            .forEach(function (k) {
                if ((C = hdr.indexOf(k)) == -1) {
                    hdr[C = hdr.length] = k
                }
                let v = js[R][k]
                let t = 'z'
                let z = ''
                if (typeof v == 'number') {
                    t = 'n'
                } else if (typeof v == 'boolean') {
                    t = 'b'
                } else if (typeof v == 'string') {
                    t = 's'
                } else if (v instanceof Date) {
                    t = 'd'
                    if (!o.cellDates) {
                        t = 'n'
                        v = datenum(v)
                    }
                    z = o.dateNF || SSF._table[14]
                }
                ws[encode_cell({ c: C, r: R + 1 })] = cell = { t: t, v: v }
                if (z) {
                    cell.z = z
                }
            })
    }
    range.e.c = hdr.length - 1
    for (C = 0; C < hdr.length; ++C) {
        ws[encode_col(C) + '1'] = { t: 's', v: hdr[C] }
    }
    ws['!ref'] = encode_range(range)
    return ws
}


export const utils = {
    encode_col,
    encode_row,
    encode_cell,
    encode_range,
    decode_col,
    decode_row,
    split_cell,
    decode_cell,
    decode_range,
    format_cell,
    get_formulae: sheet_to_formulae,
    make_csv: sheet_to_csv,
    make_json: sheet_to_json,
    make_formulae: sheet_to_formulae,
    aoa_to_sheet,
    json_to_sheet,
    table_to_sheet: parse_dom_table,
    table_to_book,
    sheet_to_csv,
    sheet_to_json,
    sheet_to_html: HTML_.from_sheet,
    sheet_to_formulae,
    sheet_to_row_object_array: sheet_to_json,
}
