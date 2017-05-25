import { DENSE } from './03_consts'
import * as SSF from './10_ssf'
import { datenum } from './20_jsutils'
import { utf8read } from './22_xmlutils'
/* [MS-XLSB] 2.4.718 BrtRowHdr */
import { new_buf } from './23_binutils'
import { buf_array, recordhopper, write_record } from './24_hoppers'
import {
    decode_cell,
    decode_range,
    encode_cell,
    encode_col,
    encode_range,
    encode_row,
    safe_decode_range
} from './27_csfutils'
import {
    BErr,
    parse_RfX,
    parse_RkNumber,
    parse_UncheckedRfX,
    parse_XLNullableWideString,
    parse_XLSBCell,
    parse_XLSBCodeName,
    parse_XLWideString,
    parse_Xnum,
    write_BrtColor,
    write_RelID,
    write_RkNumber,
    write_UncheckedRfX,
    write_XLSBCell,
    write_XLSBCodeName,
    write_XLWideString,
    write_Xnum
} from './28_binstructs'
import { add_rels, RELS } from './31_rels'
import { crypto_CreatePasswordVerifier_Method1 } from './44_offcrypto'
import { find_mdw_colw, process_col, pt2px, px2pt } from './45_styutils'
import { stringify_formula } from './62_fxls'
import { parse_XLSBArrayParsedFormula, parse_XLSBCellParsedFormula, parse_XLSBSharedParsedFormula } from './63_fbin'
import { col_obj_w, default_margins, get_cell_style, get_sst_id, safe_format, strs } from './66_wscommon'

export function parse_BrtRowHdr(data, length) {
    const z = {}
    const tgt = data.l + length
    z.r = data.read_shift(4)
    data.l += 4 // TODO: ixfe
    const miyRw = data.read_shift(2)
    data.l += 1 // TODO: top/bot padding
    const flags = data.read_shift(1)
    data.l = tgt
    if (flags & 0x10) {
        z.hidden = true
    }
    if (flags & 0x20) {
        z.hpt = miyRw / 20
    }
    return z
}

export function write_BrtRowHdr(R: number, range, ws) {
    const o = new_buf(17 + 8 * 16)
    const row = (ws['!rows'] || [])[R] || {}
    o.write_shift(4, R)

    o.write_shift(4, 0)
    /* TODO: ixfe */

    let miyRw = 0x0140
    if (row.hpx) {
        miyRw = px2pt(row.hpx) * 20
    } else if (row.hpt) {
        miyRw = row.hpt * 20
    }
    o.write_shift(2, miyRw)

    o.write_shift(1, 0)
    /* top/bot padding */

    let flags = 0x0
    if (row.hidden) {
        flags |= 0x10
    }
    if (row.hpx || row.hpt) {
        flags |= 0x20
    }
    o.write_shift(1, flags)

    o.write_shift(1, 0)
    /* phonetic guide */

    /* [MS-XLSB] 2.5.8 BrtColSpan explains the mechanism */
    let ncolspan = 0

    const lcs = o.l
    o.l += 4

    const caddr = { r: R, c: 0 }
    for (let i = 0; i < 16; ++i) {
        if (range.s.c > i + 1 << 10 || range.e.c < i << 10) {
            continue
        }
        let first = -1
        let last = -1
        for (let j = i << 10; j < i + 1 << 10; ++j) {
            caddr.c = j
            const cell = Array.isArray(ws) ? (ws[caddr.r] || [])[caddr.c] : ws[encode_cell(caddr)]
            if (cell) {
                if (first < 0) {
                    first = j
                }
                last = j
            }
        }
        if (first < 0) {
            continue
        }
        ++ncolspan
        o.write_shift(4, first)
        o.write_shift(4, last)
    }

    const l = o.l
    o.l = lcs
    o.write_shift(4, ncolspan)
    o.l = l

    return o.length > o.l ? o.slice(0, o.l) : o
}
export function write_row_header(ba, ws, range, R) {
    const o = write_BrtRowHdr(R, range, ws)
    if (o.length > 17) {
        write_record(ba, 'BrtRowHdr', o)
    }
}

/* [MS-XLSB] 2.4.812 BrtWsDim */
export const parse_BrtWsDim = parse_UncheckedRfX
export const write_BrtWsDim = write_UncheckedRfX

/* [MS-XLSB] 2.4.813 BrtWsFmtInfo */
//function write_BrtWsFmtInfo(ws, o) { }

/* [MS-XLSB] 2.4.815 BrtWsProp */
export function parse_BrtWsProp(data, length) {
    const z = {}
    /* TODO: pull flags */
    data.l += 19
    z.name = parse_XLSBCodeName(data, length - 19)
    return z
}
export function write_BrtWsProp(str, o?) {
    if (o == null) {
        o = new_buf(84 + 4 * str.length)
    }
    for (let i = 0; i < 3; ++i) {
        o.write_shift(1, 0)
    }
    write_BrtColor({ auto: 1 }, o)
    o.write_shift(-4, -1)
    o.write_shift(-4, -1)
    write_XLSBCodeName(str, o)
    return o.slice(0, o.l)
}

/* [MS-XLSB] 2.4.303 BrtCellBlank */
export function parse_BrtCellBlank(data, length) {
    const cell = parse_XLSBCell(data)
    return [cell]
}
export function write_BrtCellBlank(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(8)
    }
    return write_XLSBCell(ncell, o)
}

/* [MS-XLSB] 2.4.304 BrtCellBool */
export function parse_BrtCellBool(data, length) {
    const cell = parse_XLSBCell(data)
    const fBool = data.read_shift(1)
    return [cell, fBool, 'b']
}
export function write_BrtCellBool(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(9)
    }
    write_XLSBCell(ncell, o)
    o.write_shift(1, cell.v ? 1 : 0)
    return o
}

/* [MS-XLSB] 2.4.305 BrtCellError */
export function parse_BrtCellError(data, length) {
    const cell = parse_XLSBCell(data)
    const bError = data.read_shift(1)
    return [cell, bError, 'e']
}

/* [MS-XLSB] 2.4.308 BrtCellIsst */
export function parse_BrtCellIsst(data, length) {
    const cell = parse_XLSBCell(data)
    const isst = data.read_shift(4)
    return [cell, isst, 's']
}

export function write_BrtCellIsst(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(12)
    }
    write_XLSBCell(ncell, o)
    o.write_shift(4, ncell.v)
    return o
}

/* [MS-XLSB] 2.4.310 BrtCellReal */
export function parse_BrtCellReal(data, length) {
    const cell = parse_XLSBCell(data)
    const value = parse_Xnum(data)
    return [cell, value, 'n']
}

export function write_BrtCellReal(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(16)
    }
    write_XLSBCell(ncell, o)
    write_Xnum(cell.v, o)
    return o
}

/* [MS-XLSB] 2.4.311 BrtCellRk */
export function parse_BrtCellRk(data, length) {
    const cell = parse_XLSBCell(data)
    const value = parse_RkNumber(data)
    return [cell, value, 'n']
}

export function write_BrtCellRk(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(12)
    }
    write_XLSBCell(ncell, o)
    write_RkNumber(cell.v, o)
    return o
}

/* [MS-XLSB] 2.4.314 BrtCellSt */
export function parse_BrtCellSt(data, length) {
    const cell = parse_XLSBCell(data)
    const value = parse_XLWideString(data)
    return [cell, value, 'str']
}
export function write_BrtCellSt(cell, ncell, o?) {
    if (o == null) {
        o = new_buf(12 + 4 * cell.v.length)
    }
    write_XLSBCell(ncell, o)
    write_XLWideString(cell.v, o)
    return o.length > o.l ? o.slice(0, o.l) : o
}

/* [MS-XLSB] 2.4.647 BrtFmlaBool */
export function parse_BrtFmlaBool(data, length, opts) {
    const end = data.l + length
    const cell = parse_XLSBCell(data)
    cell.r = opts['!row']
    const value = data.read_shift(1)
    const o = [cell, value, 'b']
    if (opts.cellFormula) {
        data.l += 2
        const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts)
        o[3] = stringify_formula(formula, null /*range*/, cell, opts.supbooks, opts)
        /* TODO */
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.648 BrtFmlaError */
export function parse_BrtFmlaError(data, length, opts) {
    const end = data.l + length
    const cell = parse_XLSBCell(data)
    cell.r = opts['!row']
    const value = data.read_shift(1)
    const o = [cell, value, 'e']
    if (opts.cellFormula) {
        data.l += 2
        const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts)
        o[3] = stringify_formula(formula, null /*range*/, cell, opts.supbooks, opts)
        /* TODO */
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.649 BrtFmlaNum */
export function parse_BrtFmlaNum(data, length, opts) {
    const end = data.l + length
    const cell = parse_XLSBCell(data)
    cell.r = opts['!row']
    const value = parse_Xnum(data)
    const o = [cell, value, 'n']
    if (opts.cellFormula) {
        data.l += 2
        const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts)
        o[3] = stringify_formula(formula, null /*range*/, cell, opts.supbooks, opts)
        /* TODO */
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.650 BrtFmlaString */
export function parse_BrtFmlaString(data, length, opts) {
    const end = data.l + length
    const cell = parse_XLSBCell(data)
    cell.r = opts['!row']
    const value = parse_XLWideString(data)
    const o = [cell, value, 'str']
    if (opts.cellFormula) {
        data.l += 2
        const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts)
        o[3] = stringify_formula(formula, null /*range*/, cell, opts.supbooks, opts)
        /* TODO */
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.676 BrtMergeCell */
export const parse_BrtMergeCell = parse_UncheckedRfX
export const write_BrtMergeCell = write_UncheckedRfX
/* [MS-XLSB] 2.4.108 BrtBeginMergeCells */
export function write_BrtBeginMergeCells(cnt, o?) {
    if (o == null) {
        o = new_buf(4)
    }
    o.write_shift(4, cnt)
    return o
}

/* [MS-XLSB] 2.4.656 BrtHLink */
export function parse_BrtHLink(data, length, opts) {
    const end = data.l + length
    const rfx = parse_UncheckedRfX(data, 16)
    const relId = parse_XLNullableWideString(data)
    const loc = parse_XLWideString(data)
    const tooltip = parse_XLWideString(data)
    const display = parse_XLWideString(data)
    data.l = end
    return { rfx, relId, loc, Tooltip: tooltip, display }
}

export function write_BrtHLink(l, rId, o?) {
    if (o == null) {
        o = new_buf(50 + 4 * l[1].Target.length)
    }
    write_UncheckedRfX({ s: decode_cell(l[0]), e: decode_cell(l[0]) }, o)
    write_RelID(`rId${rId}`, o)
    const locidx = l[1].Target.indexOf('#')
    const location = locidx == -1 ? '' : l[1].Target.substr(locidx + 1)
    write_XLWideString(location || '', o)
    write_XLWideString(l[1].Tooltip || '', o)
    write_XLWideString('', o)
    return o.slice(0, o.l)
}

/* [MS-XLSB] 2.4.6 BrtArrFmla */
export function parse_BrtArrFmla(data, length, opts) {
    const end = data.l + length
    const rfx = parse_RfX(data, 16)
    const fAlwaysCalc = data.read_shift(1)
    const o = [rfx]
    o[2] = fAlwaysCalc
    if (opts.cellFormula) {
        const formula = parse_XLSBArrayParsedFormula(data, end - data.l, opts)
        o[1] = formula
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.742 BrtShrFmla */
export function parse_BrtShrFmla(data, length, opts) {
    const end = data.l + length
    const rfx = parse_UncheckedRfX(data, 16)
    const o = [rfx]
    if (opts.cellFormula) {
        const formula = parse_XLSBSharedParsedFormula(data, end - data.l, opts)
        o[1] = formula
        data.l = end
    } else {
        data.l = end
    }
    return o
}

/* [MS-XLSB] 2.4.323 BrtColInfo */
/* TODO: once XLS ColInfo is set, combine the functions */
export function write_BrtColInfo(C: number, col, o?) {
    if (o == null) {
        o = new_buf(18)
    }
    const p = col_obj_w(C, col)
    o.write_shift(-4, C)
    o.write_shift(-4, C)
    o.write_shift(4, (p.width || 10) * 256)
    o.write_shift(4, 0/*ixfe*/) // style
    let flags = 0
    if (col.hidden) {
        flags |= 0x01
    }
    if (typeof p.width == 'number') {
        flags |= 0x02
    }
    o.write_shift(1, flags) // bit flag
    o.write_shift(1, 0) // bit flag
    return o
}

/* [MS-XLSB] 2.4.672 BrtMargins */
export function parse_BrtMargins(data, length, opts) {
    return {
        left: parse_Xnum(data, 8),
        right: parse_Xnum(data, 8),
        top: parse_Xnum(data, 8),
        bottom: parse_Xnum(data, 8),
        header: parse_Xnum(data, 8),
        footer: parse_Xnum(data, 8),
    }
}
export function write_BrtMargins(margins, o?) {
    if (o == null) {
        o = new_buf(6 * 8)
    }
    default_margins(margins)
    write_Xnum(margins.left, o)
    write_Xnum(margins.right, o)
    write_Xnum(margins.top, o)
    write_Xnum(margins.bottom, o)
    write_Xnum(margins.header, o)
    write_Xnum(margins.footer, o)
    return o
}

/* [MS-XLSB] 2.4.292 BrtBeginWsView */
function write_BrtBeginWsView(ws, o) {
    if (o == null) {
        o = new_buf(30)
    }
    o.write_shift(2, 924) // bit flag
    o.write_shift(4, 0)
    o.write_shift(4, 0) // view first row
    o.write_shift(4, 0) // view first col
    o.write_shift(1, 0) // gridline color ICV
    o.write_shift(1, 0)
    o.write_shift(2, 0)
    o.write_shift(2, 100) // zoom scale
    o.write_shift(2, 0)
    o.write_shift(2, 0)
    o.write_shift(2, 0)
    o.write_shift(4, 0) // workbook view id
    return o
}

/* [MS-XLSB] 2.4.740 BrtSheetProtection */
export function write_BrtSheetProtection(sp, o?) {
    if (o == null) {
        o = new_buf(16 * 4 + 2)
    }
    o.write_shift(2, sp.password ? crypto_CreatePasswordVerifier_Method1(sp.password) : 0)
    o.write_shift(4, 1); // this record should not be written if no protection
    [
        ['objects', false], // fObjects
        ['scenarios', false], // fScenarios
        ['formatCells', true], // fFormatCells
        ['formatColumns', true], // fFormatColumns
        ['formatRows', true], // fFormatRows
        ['insertColumns', true], // fInsertColumns
        ['insertRows', true], // fInsertRows
        ['insertHyperlinks', true], // fInsertHyperlinks
        ['deleteColumns', true], // fDeleteColumns
        ['deleteRows', true], // fDeleteRows
        ['selectLockedCells', false], // fSelLockedCells
        ['sort', true], // fSort
        ['autoFilter', true], // fAutoFilter
        ['pivotTables', true], // fPivotTables
        ['selectUnlockedCells', false] // fSelUnlockedCells
    ].forEach(function (n) {
        /*:: if(o == null) throw "unreachable"; */
        if (n[1]) {
            o.write_shift(4, sp[n[0]] != null && !sp[n[0]] ? 1 : 0)
        } else {
            o.write_shift(4, sp[n[0]] != null && sp[n[0]] ? 0 : 1)
        }
    })
    return o
}

/* [MS-XLSB] 2.1.7.61 Worksheet */
export function parse_ws_bin(data, _opts, rels, wb, themes, styles): Worksheet {
    if (!data) {
        return data
    }
    const opts = _opts || {}
    if (!rels) {
        rels = { '!id': {} }
    }
    if (DENSE != null && opts.dense == null) {
        opts.dense = DENSE
    }
    const s: Worksheet = opts.dense ? [] : {}

    let ref
    const refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } }

    let pass = false
    let end = false
    let row
    let p
    let cf
    let R
    let C
    let addr
    let sstr
    let rr
    let cell: Cell
    const mergecells = []
    opts.biff = 12
    opts['!row'] = 0

    let ai = 0
    let af = false

    const array_formulae = []
    const shared_formulae = {}
    const supbooks = [[]]

    supbooks.sharedf = shared_formulae
    supbooks.arrayf = array_formulae
    supbooks.SheetNames = wb.SheetNames || wb.Sheets.map(function (x) {
        return x.name
    })
    opts.supbooks = supbooks
    for (let i = 0; i < wb.Names.length; ++i) {
        supbooks[0][i + 1] = wb.Names[i]
    }

    const colinfo = []
    const rowinfo = []
    const defwidth = 0 // twips / MDW respectively
    const defheight = 0
    let seencol = false

    recordhopper(data, function ws_parse(val, R_n, RT) {
        if (end) {
            return
        }
        switch (RT) {
            case 0x0094:
                /* 'BrtWsDim' */
                ref = val
                break
            case 0x0000:
                /* 'BrtRowHdr' */
                row = val
                if (opts.sheetRows && opts.sheetRows <= row.r) {
                    end = true
                }
                rr = encode_row(R = row.r)
                opts['!row'] = row.r
                if (val.hidden || val.hpt) {
                    if (val.hpt) {
                        val.hpx = pt2px(val.hpt)
                    }
                    rowinfo[val.r] = val
                }
                break

            case 0x0002: /* 'BrtCellRk' */
            case 0x0003: /* 'BrtCellError' */
            case 0x0004: /* 'BrtCellBool' */
            case 0x0005: /* 'BrtCellReal' */
            case 0x0006: /* 'BrtCellSt' */
            case 0x0007: /* 'BrtCellIsst' */
            case 0x0008: /* 'BrtFmlaString' */
            case 0x0009: /* 'BrtFmlaNum' */
            case 0x000A: /* 'BrtFmlaBool' */
            case 0x000B:
                /* 'BrtFmlaError' */
                p = { t: val[2] }

                switch (val[2]) {
                    case 'n':
                        p.v = val[1]
                        break
                    case 's':
                        sstr = strs[val[1]]
                        p.v = sstr.t
                        p.r = sstr.r
                        break
                    case 'b':
                        p.v = !!val[1]
                        break
                    case 'e':
                        p.v = val[1]
                        if (opts.cellText !== false) {
                            p.w = BErr[p.v]
                        }
                        break
                    case 'str':
                        p.t = 's'
                        p.v = utf8read(val[1])
                        break
                }
                if (cf = styles.CellXf[val[0].iStyleRef]) {
                    safe_format(p, cf.ifmt, null, opts, themes, styles)
                }
                C = val[0].c
                if (opts.dense) {
                    if (!s[R]) {
                        s[R] = []
                    }
                    s[R][C] = p
                } else {
                    s[encode_col(C) + rr] = p
                }
                if (opts.cellFormula) {
                    af = false
                    for (ai = 0; ai < array_formulae.length; ++ai) {
                        const aii = array_formulae[ai]
                        if (row.r >= aii[0].s.r && row.r <= aii[0].e.r) {
                            if (C >= aii[0].s.c && C <= aii[0].e.c) {
                                p.F = encode_range(aii[0])
                                af = true
                            }
                        }
                    }
                    if (!af && val.length > 3) {
                        p.f = val[3]
                    }
                }
                if (refguess.s.r > row.r) {
                    refguess.s.r = row.r
                }
                if (refguess.s.c > C) {
                    refguess.s.c = C
                }
                if (refguess.e.r < row.r) {
                    refguess.e.r = row.r
                }
                if (refguess.e.c < C) {
                    refguess.e.c = C
                }
                if (opts.cellDates && cf && p.t == 'n' && SSF.is_date(SSF._table[cf.ifmt])) {
                    const _d = SSF.parse_date_code(p.v)
                    if (_d) {
                        p.t = 'd'
                        p.v = new Date(Date.UTC(_d.y, _d.m - 1, _d.d, _d.H, _d.M, _d.S, _d.u))
                    }
                }
                break

            case 0x0001:
                /* 'BrtCellBlank' */
                if (!opts.sheetStubs) {
                    break
                }
                p = { t: 'z', v: undefined }

                C = val[0].c
                if (opts.dense) {
                    if (!s[R]) {
                        s[R] = []
                    }
                    s[R][C] = p
                } else {
                    s[encode_col(C) + rr] = p
                }
                if (refguess.s.r > row.r) {
                    refguess.s.r = row.r
                }
                if (refguess.s.c > C) {
                    refguess.s.c = C
                }
                if (refguess.e.r < row.r) {
                    refguess.e.r = row.r
                }
                if (refguess.e.c < C) {
                    refguess.e.c = C
                }
                break

            case 0x00B0:
                /* 'BrtMergeCell' */
                mergecells.push(val)
                break

            case 0x01EE:
                /* 'BrtHLink' */
                const rel = rels['!id'][val.relId]
                if (rel) {
                    val.Target = rel.Target
                    if (val.loc) {
                        val.Target += `#${val.loc}`
                    }
                    val.Rel = rel
                }
                for (R = val.rfx.s.r; R <= val.rfx.e.r; ++R) {
                    for (C = val.rfx.s.c; C <= val.rfx.e.c; ++C) {
                        if (opts.dense) {
                            if (!s[R]) {
                                s[R] = []
                            }
                            if (!s[R][C]) {
                                s[R][C] = { t: 'z', v: undefined }
                            }
                            s[R][C].l = val
                        } else {
                            addr = encode_cell({ c: C, r: R })
                            if (!s[addr]) {
                                s[addr] = { t: 'z', v: undefined }
                            }
                            s[addr].l = val
                        }
                    }
                }
                break

            case 0x01AA:
                /* 'BrtArrFmla' */
                if (!opts.cellFormula) {
                    break
                }
                array_formulae.push(val)
                cell = opts.dense ? s[R][C] : s[encode_col(C) + rr]
                cell.f = stringify_formula(val[1], refguess, { r: row.r, c: C }, supbooks, opts)
                cell.F = encode_range(val[0])
                break
            case 0x01AB:
                /* 'BrtShrFmla' */
                if (!opts.cellFormula) {
                    break
                }
                shared_formulae[encode_cell(val[0].s)] = val[1]
                cell = opts.dense ? s[R][C] : s[encode_col(C) + rr]
                cell.f = stringify_formula(val[1], refguess, { r: row.r, c: C }, supbooks, opts)
                break

            /* identical to 'ColInfo' in XLS */
            case 0x003C:
                /* 'BrtColInfo' */
                if (!opts.cellStyles) {
                    break
                }
                while (val.e >= val.s) {
                    colinfo[val.e--] = { width: val.w / 256, hidden: !!(val.flags & 0x01) }
                    if (!seencol) {
                        seencol = true
                        find_mdw_colw(val.w / 256)
                    }
                    process_col(colinfo[val.e + 1])
                }
                break

            case 0x00A1:
                /* 'BrtBeginAFilter' */
                s['!autofilter'] = { ref: encode_range(val) }
                break

            case 0x01DC:
                /* 'BrtMargins' */
                s['!margins'] = val
                break

            /* case 'BrtUid' */
            case 0x00AF: /* 'BrtAFilterDateGroupItem' */
            case 0x0284: /* 'BrtActiveX' */
            case 0x0271: /* 'BrtBigName' */
            case 0x0232: /* 'BrtBkHim' */
            case 0x018C: /* 'BrtBrk' */
            case 0x0458: /* 'BrtCFIcon' */
            case 0x047A: /* 'BrtCFRuleExt' */
            case 0x01D7: /* 'BrtCFVO' */
            case 0x041A: /* 'BrtCFVO14' */
            case 0x0289: /* 'BrtCellIgnoreEC' */
            case 0x0451: /* 'BrtCellIgnoreEC14' */
            case 0x0031: /* 'BrtCellMeta' */
            case 0x024D: /* 'BrtCellSmartTagProperty' */
            case 0x025F: /* 'BrtCellWatch' */
            case 0x0234: /* 'BrtColor' */
            case 0x041F: /* 'BrtColor14' */
            case 0x00A8: /* 'BrtColorFilter' */
            case 0x00AE: /* 'BrtCustomFilter' */
            case 0x049C: /* 'BrtCustomFilter14' */
            case 0x01F3: /* 'BrtDRef' */
            case 0x0040: /* 'BrtDVal' */
            case 0x041D: /* 'BrtDVal14' */
            case 0x0226: /* 'BrtDrawing' */
            case 0x00AB: /* 'BrtDynamicFilter' */
            case 0x00A7: /* 'BrtFilter' */
            case 0x0499: /* 'BrtFilter14' */
            case 0x00A9: /* 'BrtIconFilter' */
            case 0x049D: /* 'BrtIconFilter14' */
            case 0x0227: /* 'BrtLegacyDrawing' */
            case 0x0228: /* 'BrtLegacyDrawingHF' */
            case 0x0295: /* 'BrtListPart' */
            case 0x027F: /* 'BrtOleObject' */
            case 0x01DE: /* 'BrtPageSetup' */
            case 0x0097: /* 'BrtPane' */
            case 0x0219: /* 'BrtPhoneticInfo' */
            case 0x01DD: /* 'BrtPrintOptions' */
            case 0x0218: /* 'BrtRangeProtection' */
            case 0x044F: /* 'BrtRangeProtection14' */
            case 0x02A8: /* 'BrtRangeProtectionIso' */
            case 0x0450: /* 'BrtRangeProtectionIso14' */
            case 0x0400: /* 'BrtRwDescent' */
            case 0x0098: /* 'BrtSel' */
            case 0x0297: /* 'BrtSheetCalcProp' */
            case 0x0217: /* 'BrtSheetProtection' */
            case 0x02A6: /* 'BrtSheetProtectionIso' */
            case 0x01F8: /* 'BrtSlc' */
            case 0x0413: /* 'BrtSparkline' */
            case 0x01AC: /* 'BrtTable' */
            case 0x00AA: /* 'BrtTop10Filter' */
            case 0x0032: /* 'BrtValueMeta' */
            case 0x0816: /* 'BrtWebExtension' */
            case 0x01E5: /* 'BrtWsFmtInfo' */
            case 0x0415: /* 'BrtWsFmtInfoEx14' */
            case 0x0093:
                /* 'BrtWsProp' */
                break

            case 0x0023:
                /* 'BrtFRTBegin' */
                pass = true
                break
            case 0x0024:
                /* 'BrtFRTEnd' */
                pass = false
                break
            case 0x0025:
                /* 'BrtACBegin' */
                break
            case 0x0026:
                /* 'BrtACEnd' */
                break

            default:
                if ((R_n || '').indexOf('Begin') > 0) {
                } else if ((R_n || '').indexOf('End') > 0) {
                } else if (!pass || opts.WTF) {
                    throw new Error(`Unexpected record ${RT} ${R_n}`)
                }
        }
    }, opts)

    delete opts.supbooks
    delete opts['!row']

    if (!s['!ref'] && (refguess.s.r < 2000000 || ref && (ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0))) {
        s['!ref'] = encode_range(ref || refguess)
    }
    if (opts.sheetRows && s['!ref']) {
        const tmpref = safe_decode_range(s['!ref'])
        if (opts.sheetRows < +tmpref.e.r) {
            tmpref.e.r = opts.sheetRows - 1
            if (tmpref.e.r > refguess.e.r) {
                tmpref.e.r = refguess.e.r
            }
            if (tmpref.e.r < tmpref.s.r) {
                tmpref.s.r = tmpref.e.r
            }
            if (tmpref.e.c > refguess.e.c) {
                tmpref.e.c = refguess.e.c
            }
            if (tmpref.e.c < tmpref.s.c) {
                tmpref.s.c = tmpref.e.c
            }
            s['!fullref'] = s['!ref']
            s['!ref'] = encode_range(tmpref)
        }
    }
    if (mergecells.length > 0) {
        s['!merges'] = mergecells
    }
    if (colinfo.length > 0) {
        s['!cols'] = colinfo
    }
    if (rowinfo.length > 0) {
        s['!rows'] = rowinfo
    }
    return s
}

/* TODO: something useful -- this is a stub */
function write_ws_bin_cell(ba: BufArray, cell: Cell, R: number, C: number, opts, ws: Worksheet) {
    if (cell.v === undefined) {
        return ''
    }
    let vv = ''
    let olddate = null
    switch (cell.t) {
        case 'b':
            vv = cell.v ? '1' : '0'
            break
        case 'd':
            // no BrtCellDate :(
            cell.z = cell.z || SSF._table[14]
            olddate = cell.v
            cell.v = datenum(cell.v)
            cell.t = 'n'
            break
        /* falls through */
        case 'n':
        case 'e':
            vv = `${cell.v}`
            break
        default:
            vv = cell.v
            break
    }
    const o = { r: R, c: C }

    /* TODO: cell style */
    o.s = get_cell_style(opts.cellXfs, cell, opts)
    if (cell.l) {
        ws['!links'].push([encode_cell(o), cell.l])
    }
    if (cell.c) {
        ws['!comments'].push([encode_cell(o), cell.c])
    }
    switch (cell.t) {
        case 's':
        case 'str':
            if (opts.bookSST) {
                vv = get_sst_id(opts.Strings, cell.v)
                o.t = 's'
                o.v = vv
                write_record(ba, 'BrtCellIsst', write_BrtCellIsst(cell, o))
            } else {
                o.t = 'str'
                write_record(ba, 'BrtCellSt', write_BrtCellSt(cell, o))
            }
            return
        case 'n':
            /* TODO: determine threshold for Real vs RK */
            if (cell.v == (cell.v | 0) && cell.v > -1000 && cell.v < 1000) {
                write_record(ba, 'BrtCellRk', write_BrtCellRk(cell, o))
            } else {
                write_record(ba, 'BrtCellReal', write_BrtCellReal(cell, o))
            }
            if (olddate) {
                cell.t = 'd'
                cell.v = olddate
            }
            return
        case 'b':
            o.t = 'b'
            write_record(ba, 'BrtCellBool', write_BrtCellBool(cell, o))
            return
        case 'e':
            /* TODO: error */
            o.t = 'e'
            break
    }
    write_record(ba, 'BrtCellBlank', write_BrtCellBlank(cell, o))
}

function write_CELLTABLE(ba, ws: Worksheet, idx: number, opts, wb: Workbook) {
    const range = safe_decode_range(ws['!ref'] || 'A1')
    let ref
    let rr = ''
    const cols = []
    write_record(ba, 'BrtBeginSheetData')
    const dense = Array.isArray(ws)
    for (let R = range.s.r; R <= range.e.r; ++R) {
        rr = encode_row(R)
        /* [ACCELLTABLE] */
        /* BrtRowHdr */
        write_row_header(ba, ws, range, R)
        for (let C = range.s.c; C <= range.e.c; ++C) {
            /* *16384CELL */
            if (R === range.s.r) {
                cols[C] = encode_col(C)
            }
            ref = cols[C] + rr
            const cell = dense ? (ws[R] || [])[C] : ws[ref]
            if (!cell) {
                continue
            }
            /* write cell */
            write_ws_bin_cell(ba, cell, R, C, opts, ws)
        }
    }
    write_record(ba, 'BrtEndSheetData')
}

function write_MERGECELLS(ba, ws: Worksheet) {
    if (!ws || !ws['!merges']) {
        return
    }
    write_record(ba, 'BrtBeginMergeCells', write_BrtBeginMergeCells(ws['!merges'].length))
    ws['!merges'].forEach(function (m) {
        write_record(ba, 'BrtMergeCell', write_BrtMergeCell(m))
    })
    write_record(ba, 'BrtEndMergeCells')
}

function write_COLINFOS(ba, ws: Worksheet, idx: number, opts, wb: Workbook) {
    if (!ws || !ws['!cols']) {
        return
    }
    write_record(ba, 'BrtBeginColInfos')
    ws['!cols'].forEach(function (m, i) {
        if (m) {
            write_record(ba, 'BrtColInfo', write_BrtColInfo(i, m))
        }
    })
    write_record(ba, 'BrtEndColInfos')
}

function write_HLINKS(ba, ws: Worksheet, rels) {
    /* *BrtHLink */
    ws['!links'].forEach(function (l) {
        if (!l[1].Target) {
            return
        }
        const rId = add_rels(rels, -1, l[1].Target.replace(/#.*$/, ''), RELS.HLINK)
        write_record(ba, 'BrtHLink', write_BrtHLink(l, rId))
    })
    delete ws['!links']
}
function write_LEGACYDRAWING(ba, ws: Worksheet, idx: number, rels) {
    /* [BrtLegacyDrawing] */
    if (ws['!comments'].length > 0) {
        const rId = add_rels(rels, -1, `../drawings/vmlDrawing${idx + 1}.vml`, RELS.VML)
        write_record(ba, 'BrtLegacyDrawing', write_RelID(`rId${rId}`))
        ws['!legacy'] = rId
    }
}

function write_AUTOFILTER(ba, ws) {
    if (!ws['!autofilter']) {
        return
    }
    write_record(ba, 'BrtBeginAFilter', write_UncheckedRfX(decode_range(ws['!autofilter'].ref)))
    /* *FILTERCOLUMN */
    /* [SORTSTATE] */
    /* BrtEndAFilter */
    write_record(ba, 'BrtEndAFilter')
}

function write_WSVIEWS2(ba, ws) {
    write_record(ba, 'BrtBeginWsViews')
    { /* 1*WSVIEW2 */
        /* [ACUID] */
        write_record(ba, 'BrtBeginWsView', write_BrtBeginWsView(ws))
        /* [BrtPane] */
        /* *4BrtSel */
        /* *4SXSELECT */
        /* *FRT */
        write_record(ba, 'BrtEndWsView')
    }
    /* *FRT */
    write_record(ba, 'BrtEndWsViews')
}

function write_WSFMTINFO(ba, ws) {
    /* [ACWSFMTINFO] */
    //write_record(ba, "BrtWsFmtInfo", write_BrtWsFmtInfo(ws));
}

function write_SHEETPROTECT(ba, ws) {
    if (!ws['!protect']) {
        return
    }
    /* [BrtSheetProtectionIso] */
    write_record(ba, 'BrtSheetProtection', write_BrtSheetProtection(ws['!protect']))
}

export function write_ws_bin(idx: number, opts, wb: Workbook, rels) {
    const ba = buf_array()
    const s = wb.SheetNames[idx]
    const ws = wb.Sheets[s] || {}
    const r = safe_decode_range(ws['!ref'] || 'A1')
    ws['!links'] = []
    /* passed back to write_zip and removed there */
    ws['!comments'] = []
    write_record(ba, 'BrtBeginSheet')
    write_record(ba, 'BrtWsProp', write_BrtWsProp(s))
    write_record(ba, 'BrtWsDim', write_BrtWsDim(r))
    write_WSVIEWS2(ba, ws)
    write_WSFMTINFO(ba, ws)
    write_COLINFOS(ba, ws, idx, opts, wb)
    write_CELLTABLE(ba, ws, idx, opts, wb)
    /* [BrtSheetCalcProp] */
    write_SHEETPROTECT(ba, ws)
    /* *([BrtRangeProtectionIso] BrtRangeProtection) */
    /* [SCENMAN] */
    write_AUTOFILTER(ba, ws)
    /* [SORTSTATE] */
    /* [DCON] */
    /* [USERSHVIEWS] */
    write_MERGECELLS(ba, ws)
    /* [BrtPhoneticInfo] */
    /* *CONDITIONALFORMATTING */
    /* [DVALS] */
    write_HLINKS(ba, ws, rels)
    /* [BrtPrintOptions] */
    if (ws['!margins']) {
        write_record(ba, 'BrtMargins', write_BrtMargins(ws['!margins']))
    }
    /* [BrtPageSetup] */
    /* [HEADERFOOTER] */
    /* [RWBRK] */
    /* [COLBRK] */
    /* *BrtBigName */
    /* [CELLWATCHES] */
    /* [IGNOREECS] */
    /* [SMARTTAGS] */
    /* [BrtDrawing] */
    write_LEGACYDRAWING(ba, ws, idx, rels)
    /* [BrtLegacyDrawingHF] */
    /* [BrtBkHim] */
    /* [OLEOBJECTS] */
    /* [ACTIVEXCONTROLS] */
    /* [WEBPUBITEMS] */
    /* [LISTPARTS] */
    /* FRTWORKSHEET */
    write_record(ba, 'BrtEndSheet')
    return ba.end()
}
