import * as SSF from './10_ssf'
import { datenum, parseDate } from './20_jsutils'
import { BErr } from './28_binstructs'
import { RELS } from './31_rels'
import { char2width, px2char, rgb_tint, setMDW } from './45_styutils'

export const strs = {} // shared strings
export let _ssfopts = {} // spreadsheet formatting options

export function resetSSFOpts (val) {
    _ssfopts = val
}

RELS.WS = [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet',
]

export function get_sst_id(sst /*:SST*/, str /*:string*/) /*:number*/ {
    const len = sst.length
    for (let i = 0; i < len; ++i) {
        if (sst[i].t === str) {
            sst.Count++
            return i
        }
    }
    sst[len] = { t: str }
    sst.Count++
    sst.Unique++
    return len
}

export function col_obj_w(C /*:number*/, col) {
    const p = { min: C + 1, max: C + 1 }
    /*:any*/
    /* wch (chars), wpx (pixels) */
    let wch = -1
    if (col.MDW) {
        setMDW(col.MDW)
    }
    if (col.width != null) {
        p.customWidth = 1
    } else if (col.wpx != null) {
        wch = px2char(col.wpx)
    } else if (col.wch != null) {
        wch = col.wch
    }
    if (wch > -1) {
        p.width = char2width(wch)
        p.customWidth = 1
    } else if (col.width != null) {
        p.width = col.width
    }
    if (col.hidden) {
        p.hidden = true
    }
    return p
}

export function default_margins(margins, mode?) {
    if (!margins) {
        return
    }
    let defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3]
    if (mode == 'xlml') {
        defs = [1, 1, 1, 1, 0.5, 0.5]
    }
    if (margins.left == null) {
        margins.left = defs[0]
    }
    if (margins.right == null) {
        margins.right = defs[1]
    }
    if (margins.top == null) {
        margins.top = defs[2]
    }
    if (margins.bottom == null) {
        margins.bottom = defs[3]
    }
    if (margins.header == null) {
        margins.header = defs[4]
    }
    if (margins.footer == null) {
        margins.footer = defs[5]
    }
}

export function get_cell_style(styles, cell, opts) {
    const z = opts.revssf[cell.z != null ? cell.z : 'General']
    const len = styles.length
    for (let i = 0; i != len; ++i) {
        if (styles[i].numFmtId === z) {
            return i
        }
    }
    styles[len] = {
        numFmtId: z,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        xfId: 0,
        applyNumberFormat: 1,
    }
    return len
}

export function safe_format(p, fmtid, fillid, opts, themes, styles) {
    if (p.t === 'z') {
        return
    }
    if (p.t === 'd' && typeof p.v === 'string') {
        p.v = parseDate(p.v)
    }
    try {
        if (opts.cellNF) {
            p.z = SSF._table[fmtid]
        }
    } catch (e) {
        if (opts.WTF) {
            throw e
        }
    }
    if (!opts || opts.cellText !== false) {
        try {
            if (p.t === 'e') {
                p.w = p.w || BErr[p.v]
            } else if (fmtid === 0) {
                if (p.t === 'n') {
                    if ((p.v | 0) === p.v) {
                        p.w = SSF._general_int(p.v, _ssfopts)
                    } else {
                        p.w = SSF._general_num(p.v, _ssfopts)
                    }
                } else if (p.t === 'd') {
                    const dd = datenum(p.v)
                    if ((dd | 0) === dd) {
                        p.w = SSF._general_int(dd, _ssfopts)
                    } else {
                        p.w = SSF._general_num(dd, _ssfopts)
                    }
                } else if (p.v === undefined) {
                    return ''
                } else {
                    p.w = SSF._general(p.v, _ssfopts)
                }
            } else if (p.t === 'd') {
                p.w = SSF.format(fmtid, datenum(p.v), _ssfopts)
            } else {
                p.w = SSF.format(fmtid, p.v, _ssfopts)
            }
        } catch (e) {
            if (opts.WTF) {
                throw e
            }
        }
    }
    if (fillid) {
        try {
            p.s = styles.Fills[fillid]
            if (p.s.fgColor && p.s.fgColor.theme) {
                p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0)
                if (opts.WTF) {
                    p.s.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fgColor.theme].rgb
                }
            }
            if (p.s.bgColor && p.s.bgColor.theme) {
                p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0)
                if (opts.WTF) {
                    p.s.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.bgColor.theme].rgb
                }
            }
        } catch (e) {
            if (opts.WTF) {
                throw e
            }
        }
    }
}
