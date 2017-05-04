import { DENSE } from './03_consts'
import { SSF } from './10_ssf'
import { datenum, dup, parseDate } from './20_jsutils'
import {
    escapehtml,
    escapexml,
    matchtag,
    parsexmlbool,
    parsexmltag,
    unescapexml,
    utf8read,
    writetag,
    writextag,
    XML_HEADER,
    XMLNS
} from './22_xmlutils'
import { decode_cell, encode_cell, encode_col, encode_range, encode_row, safe_decode_range } from './27_csfutils'
import { BErr, RBErr } from './28_binstructs'
import { add_rels, RELS } from './31_rels'
import { parse_si } from './42_sstxml'
import { crypto_CreatePasswordVerifier_Method1 } from './44_offcrypto'
import { find_mdw_colw, process_col, pt2px, px2pt } from './45_styutils'
import { shift_formula_xlsx } from './61_fcommon'
import { col_obj_w, default_margins, get_cell_style, get_sst_id, safe_format, strs } from './66_wscommon'

function parse_ws_xml_dim(ws, s) {
    const d = safe_decode_range(s)
    if (d.s.r <= d.e.r
        && d.s.c <= d.e.c
        && d.s.r >= 0
        && d.s.c >= 0
    ) {
        ws['!ref'] = encode_range(d)
    }
}

const mergecregex = /<(?:\w:)?mergeCell ref="[A-Z0-9:]+"\s*[\/]?>/g
const sheetdataregex = /<(?:\w+:)?sheetData>([^\u2603]*)<\/(?:\w+:)?sheetData>/
const hlinkregex = /<(?:\w:)?hyperlink [^>]*>/mg
const dimregex = /"(\w*:\w*)"/
const colregex = /<(?:\w:)?col[^>]*[\/]?>/g
const afregex = /<(?:\w:)?autoFilter[^>]*([\/]|>([^\u2603]*)<\/(?:\w:)?autoFilter)>/g
const marginregex = /<(?:\w:)?pageMargins[^>]*\/>/g

/* 18.3 Worksheets */
export function parse_ws_xml(data /*:?string*/, opts, rels, wb, themes, styles) /*:Worksheet*/ {
    if (!data) {
        return data
    }
    if (DENSE != null && opts.dense == null) {
        opts.dense = DENSE
    }

    /* 18.3.1.99 worksheet CT_Worksheet */
    const s = opts.dense ? [] /*:any*/ : {}
    /*:any*/
    const refguess /*:Range*/ = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } }
    /*:any*/

    let data1 = ''
    let data2 = ''
    const mtch = data.match(sheetdataregex)
    if (mtch) {
        data1 = data.substr(0, mtch.index)
        data2 = data.substr(mtch.index + mtch[0].length)
    } else {
        data1 = data2 = data
    }

    /* 18.3.1.35 dimension CT_SheetDimension ? */
    // $FlowIgnore
    let ridx = (data1.match(/<(?:\w*:)?dimension/) || { index: -1 }).index
    if (ridx > 0) {
        const ref = data1.substr(ridx, 50).match(dimregex)
        if (ref) {
            parse_ws_xml_dim(s, ref[1])
        }
    }

    /* 18.3.1.17 cols CT_Cols */
    const columns = []
    if (opts.cellStyles) {
        /* 18.3.1.13 col CT_Col */
        const cols = data1.match(colregex)
        if (cols) {
            parse_ws_xml_cols(columns, cols)
        }
    }

    /* 18.3.1.80 sheetData CT_SheetData ? */
    if (mtch) {
        parse_ws_xml_data(mtch[1], s, opts, refguess, themes, styles)
    }

    /* 18.3.1.2  autoFilter CT_AutoFilter */
    const afilter = data2.match(afregex)
    if (afilter) {
        s['!autofilter'] = parse_ws_xml_autofilter(afilter[0])
    }

    /* 18.3.1.55 mergeCells CT_MergeCells */
    const mergecells = []
    const merges = data2.match(mergecregex)
    if (merges) {
        for (ridx = 0; ridx != merges.length; ++ridx) {
            mergecells[ridx] = safe_decode_range(merges[ridx].substr(merges[ridx].indexOf('"') + 1))
        }
    }

    /* 18.3.1.48 hyperlinks CT_Hyperlinks */
    const hlink = data2.match(hlinkregex)
    if (hlink) {
        parse_ws_xml_hlinks(s, hlink, rels)
    }

    /* 18.3.1.62 pageMargins CT_PageMargins */
    const margins = data2.match(marginregex)
    if (margins) {
        s['!margins'] = parse_ws_xml_margins(parsexmltag(margins[0]))
    }

    if (!s['!ref'] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) {
        s['!ref'] = encode_range(refguess)
    }
    if (opts.sheetRows > 0 && s['!ref']) {
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
    if (columns.length > 0) {
        s['!cols'] = columns
    }
    return s
}

function write_ws_xml_merges(merges) {
    if (merges.length == 0) {
        return ''
    }
    let o = `<mergeCells count="${merges.length}">`
    for (let i = 0; i != merges.length; ++i) {
        o += `<mergeCell ref="${encode_range(merges[i])}"/>`
    }
    return `${o}</mergeCells>`
}

/* 18.3.1.85 sheetPr CT_SheetProtection */
function write_ws_xml_protection(sp) /*:string*/ {
    // algorithmName, hashValue, saltValue, spinCountpassword
    const o = { sheet: 1 }
    /*:any*/
    const deffalse = ['objects', 'scenarios', 'selectLockedCells', 'selectUnlockedCells']
    const deftrue = [
        'formatColumns',
        'formatRows',
        'formatCells',
        'insertColumns',
        'insertRows',
        'insertHyperlinks',
        'deleteColumns',
        'deleteRows',
        'sort',
        'autoFilter',
        'pivotTables',
    ]
    deffalse.forEach(function (n) {
        if (sp[n] != null && sp[n]) {
            o[n] = '1'
        }
    })
    deftrue.forEach(function (n) {
        if (sp[n] != null && !sp[n]) {
            o[n] = '0'
        }
    })
    /* TODO: algorithm */
    if (sp.password) {
        o.password = crypto_CreatePasswordVerifier_Method1(sp.password).toString(16).toUpperCase()
    }
    return writextag('sheetProtection', null, o)
}

function parse_ws_xml_hlinks(s, data /*:Array<string>*/, rels) {
    const dense = Array.isArray(s)
    for (let i = 0; i != data.length; ++i) {
        const val = parsexmltag(data[i], true)
        if (!val.ref) {
            return
        }
        let rel = rels ? rels['!id'][val.id] : null
        if (rel) {
            val.Target = rel.Target
            if (val.location) {
                val.Target += `#${val.location}`
            }
            val.Rel = rel
        } else {
            val.Target = val.location
            rel = { Target: val.location, TargetMode: 'Internal' }
            val.Rel = rel
        }
        if (val.tooltip) {
            val.Tooltip = val.tooltip
            delete val.tooltip
        }
        const rng = safe_decode_range(val.ref)
        for (let R = rng.s.r; R <= rng.e.r; ++R) {
            for (let C = rng.s.c; C <= rng.e.c; ++C) {
                const addr = encode_cell({ c: C, r: R })
                if (dense) {
                    if (!s[R]) {
                        s[R] = []
                    }
                    if (!s[R][C]) {
                        s[R][C] = { t: 'z', v: undefined }
                    }
                    s[R][C].l = val
                } else {
                    if (!s[addr]) {
                        s[addr] = { t: 'z', v: undefined }
                    }
                    s[addr].l = val
                }
            }
        }
    }
}

function parse_ws_xml_margins(margin) {
    const o = {};
    ['left', 'right', 'top', 'bottom', 'header', 'footer'].forEach(function (k) {
        if (margin[k]) {
            o[k] = parseFloat(margin[k])
        }
    })
    return o
}
function write_ws_xml_margins(margin) {
    default_margins(margin)
    return writextag('pageMargins', null, margin)
}

function parse_ws_xml_cols(columns, cols) {
    let seencol = false
    for (let coli = 0; coli != cols.length; ++coli) {
        const coll = parsexmltag(cols[coli], true)
        if (coll.hidden) {
            coll.hidden = parsexmlbool(coll.hidden)
        }
        let colm = parseInt(coll.min, 10) - 1
        const colM = parseInt(coll.max, 10) - 1
        delete coll.min
        delete coll.max
        coll.width = +coll.width
        if (!seencol && coll.width) {
            seencol = true
            find_mdw_colw(coll.width)
        }
        process_col(coll)
        while (colm <= colM) {
            columns[colm++] = dup(coll)
        }
    }
}

function write_ws_xml_cols(ws, cols) /*:string*/ {
    const o = ['<cols>']
    let col
    let width
    for (let i = 0; i != cols.length; ++i) {
        if (!(col = cols[i])) {
            continue
        }
        o[o.length] = writextag('col', null, col_obj_w(i, col))
    }
    o[o.length] = '</cols>'
    return o.join('')
}

function parse_ws_xml_autofilter(data) {
    const o = { ref: (data.match(/ref="([^"]*)"/) || [])[1] }
    return o
}
function write_ws_xml_autofilter(data) /*:string*/ {
    return writextag('autoFilter', null, { ref: data.ref })
}

/* 18.3.1.88 sheetViews CT_SheetViews */
/* 18.3.1.87 sheetView CT_SheetView */
function write_ws_xml_sheetviews(ws, opts, idx, wb)/*:string*/ {
    return writextag('sheetViews', writextag('sheetView', null, { workbookViewId: '0' }), {})
}

function write_ws_xml_cell(cell, ref, ws, opts, idx, wb) {
    if (cell.v === undefined && cell.f === undefined || cell.t === 'z') {
        return ''
    }
    let vv = ''
    const oldt = cell.t
    const oldv = cell.v
    switch (cell.t) {
        case 'b':
            vv = cell.v ? '1' : '0'
            break
        case 'n':
            vv = `${cell.v}`
            break
        case 'e':
            vv = BErr[cell.v]
            break
        case 'd':
            if (opts.cellDates) {
                vv = parseDate(cell.v).toISOString()
            } else {
                cell.t = 'n'
                vv = `${cell.v = datenum(parseDate(cell.v))}`
                if (typeof cell.z === 'undefined') {
                    cell.z = SSF._table[14]
                }
            }
            break
        default:
            vv = cell.v
            break
    }
    let v = writetag('v', escapexml(vv))
    const o = { r: ref }
    /*:any*/
    /* TODO: cell style */
    const os = get_cell_style(opts.cellXfs, cell, opts)
    if (os !== 0) {
        o.s = os
    }
    switch (cell.t) {
        case 'n':
            break
        case 'd':
            o.t = 'd'
            break
        case 'b':
            o.t = 'b'
            break
        case 'e':
            o.t = 'e'
            break
        default:
            if (cell.v == null) {
                delete cell.t
                break
            }
            if (opts.bookSST) {
                v = writetag('v', `${get_sst_id(opts.Strings, cell.v)}`)
                o.t = 's'
                break
            }
            o.t = 'str'
            break
    }
    if (cell.t != oldt) {
        cell.t = oldt
        cell.v = oldv
    }
    if (cell.f) {
        const ff = cell.F && cell.F.substr(0, ref.length) == ref ? { t: 'array', ref: cell.F } : null
        v = writextag('f', escapexml(cell.f), ff) + (cell.v != null ? v : '')
    }
    if (cell.l) {
        ws['!links'].push([ref, cell.l])
    }
    if (cell.c) {
        ws['!comments'].push([ref, cell.c])
    }
    return writextag('c', v, o)
}

const parse_ws_xml_data = function parse_ws_xml_data_factory() {
    const cellregex = /<(?:\w+:)?c[ >]/
    const rowregex = /<\/(?:\w+:)?row>/
    const rregex = /r=["']([^"']*)["']/
    const isregex = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/
    const refregex = /ref=["']([^"']*)["']/
    const match_v = matchtag('v')
    const match_f = matchtag('f')

    return function parse_ws_xml_data(sdata, s, opts, guess, themes, styles) {
        let ri = 0
        let x = ''
        let cells = []
        let cref = []
        let idx = 0
        let i = 0
        let cc = 0
        let d = ''
        let p
        /*:any*/
        let tag
        let tagr = 0
        let tagc = 0
        let sstr
        let ftag
        let fmtid = 0
        let fillid = 0
        const do_format = Array.isArray(styles.CellXf)
        let cf
        const arrayf = []
        const sharedf = []
        const dense = Array.isArray(s)
        const rows = []
        let rowobj = {}
        let rowrite = false
        for (let marr = sdata.split(rowregex), mt = 0, marrlen = marr.length; mt != marrlen; ++mt) {
            x = marr[mt].trim()
            const xlen = x.length
            if (xlen === 0) {
                continue
            }

            /* 18.3.1.73 row CT_Row */
            for (ri = 0; ri < xlen; ++ri) {
                if (x.charCodeAt(ri) === 62) {
                    break
                }
            }
            ++ri
            tag = parsexmltag(x.substr(0, ri), true)
            tagr = tag.r != null ? parseInt(tag.r, 10) : tagr + 1
            tagc = -1
            if (opts.sheetRows && opts.sheetRows < tagr) {
                continue
            }
            if (guess.s.r > tagr - 1) {
                guess.s.r = tagr - 1
            }
            if (guess.e.r < tagr - 1) {
                guess.e.r = tagr - 1
            }

            if (opts && opts.cellStyles) {
                rowobj = {}
                rowrite = false
                if (tag.ht) {
                    rowrite = true
                    rowobj.hpt = parseFloat(tag.ht)
                    rowobj.hpx = pt2px(rowobj.hpt)
                }
                if (tag.hidden == '1') {
                    rowrite = true
                    rowobj.hidden = true
                }
                if (rowrite) {
                    rows[tagr - 1] = rowobj
                }
            }

            /* 18.3.1.4 c CT_Cell */
            cells = x.substr(ri).split(cellregex)
            for (ri = 0; ri != cells.length; ++ri) {
                x = cells[ri].trim()
                if (x.length === 0) {
                    continue
                }
                cref = x.match(rregex)
                idx = ri
                i = 0
                cc = 0
                x = `<c ${x.substr(0, 1) == '<' ? '>' : ''}${x}`
                if (cref != null && cref.length === 2) {
                    idx = 0
                    d = cref[1]
                    for (i = 0; i != d.length; ++i) {
                        if ((cc = d.charCodeAt(i) - 64) < 1 || cc > 26) {
                            break
                        }
                        idx = 26 * idx + cc
                    }
                    --idx
                    tagc = idx
                } else {
                    ++tagc
                }
                for (i = 0; i != x.length; ++i) {
                    if (x.charCodeAt(i) === 62) {
                        break
                    }
                }
                ++i
                tag = parsexmltag(x.substr(0, i), true)
                if (!tag.r) {
                    tag.r = encode_cell({ r: tagr - 1, c: tagc })
                }
                d = x.substr(i)
                p = { t: '' }
                /*:any*/

                if ((cref = d.match(match_v)) != null && /*::cref != null && */cref[1] !== '') {
                    p.v = unescapexml(cref[1])
                }
                if (opts.cellFormula) {
                    if ((cref = d.match(match_f)) != null && /*::cref != null && */cref[1] !== '') {
                        /* TODO: match against XLSXFutureFunctions */
                        p.f = unescapexml(utf8read(cref[1])).replace(/_xlfn\./, '')

                        if (cref[0].includes('t="array"')) {
                            p.F = (d.match(refregex) || [])[1]
                            if (p.F.includes(':')) {
                                arrayf.push([safe_decode_range(p.F), p.F])
                            }
                        } else if (cref[0].includes('t="shared"')) {
                            // TODO: parse formula
                            ftag = parsexmltag(cref[0])
                            sharedf[parseInt(ftag.si, 10)] = [ftag, unescapexml(utf8read(cref[1]))]
                        }
                    } else if (cref = d.match(/<f[^>]*\/>/)) {
                        ftag = parsexmltag(cref[0])
                        if (sharedf[ftag.si]) {
                            p.f = shift_formula_xlsx(sharedf[ftag.si][1], sharedf[ftag.si][0].ref, tag.r)
                        }
                    }
                    /* TODO: factor out contains logic */
                    const _tag = decode_cell(tag.r)
                    for (i = 0; i < arrayf.length; ++i) {
                        if (_tag.r >= arrayf[i][0].s.r && _tag.r <= arrayf[i][0].e.r) {
                            if (_tag.c >= arrayf[i][0].s.c && _tag.c <= arrayf[i][0].e.c) {
                                p.F = arrayf[i][1]
                            }
                        }
                    }
                }

                if (tag.t == null && p.v === undefined) {
                    if (!opts.sheetStubs) {
                        continue
                    }
                    p.t = 'z'
                } else {
                    p.t = tag.t || 'n'
                }
                if (guess.s.c > idx) {
                    guess.s.c = idx
                }
                if (guess.e.c < idx) {
                    guess.e.c = idx
                }
                /* 18.18.11 t ST_CellType */
                switch (p.t) {
                    case 'n':
                        p.v = parseFloat(p.v)
                        break
                    case 's':
                        sstr = strs[parseInt(p.v, 10)]
                        if (typeof p.v == 'undefined') {
                            if (!opts.sheetStubs) {
                                continue
                            }
                            p.t = 'z'
                        }
                        p.v = sstr.t
                        p.r = sstr.r
                        if (opts.cellHTML) {
                            p.h = sstr.h
                        }
                        break
                    case 'str':
                        p.t = 's'
                        p.v = p.v != null ? utf8read(p.v) : ''
                        if (opts.cellHTML) {
                            p.h = escapehtml(p.v)
                        }
                        break
                    case 'inlineStr':
                        cref = d.match(isregex)
                        p.t = 's'
                        if (cref != null && (sstr = parse_si(cref[1]))) {
                            p.v = sstr.t
                        } else {
                            p.v = ''
                        }
                        break
                    case 'b':
                        p.v = parsexmlbool(p.v)
                        break
                    case 'd':
                        if (!opts.cellDates) {
                            p.v = datenum(parseDate(p.v))
                            p.t = 'n'
                        }
                        break
                    /* error string in .w, number in .v */
                    case 'e':
                        if (!opts || opts.cellText !== false) {
                            p.w = p.v
                        }
                        p.v = RBErr[p.v]
                        break
                }
                /* formatting */
                fmtid = fillid = 0
                if (do_format && tag.s !== undefined) {
                    cf = styles.CellXf[tag.s]
                    if (cf != null) {
                        if (cf.numFmtId != null) {
                            fmtid = cf.numFmtId
                        }
                        if (opts.cellStyles && cf.fillId != null) {
                            fillid = cf.fillId
                        }
                    }
                }
                safe_format(p, fmtid, fillid, opts, themes, styles)
                if (opts.cellDates && do_format && p.t == 'n' && SSF.is_date(SSF._table[fmtid])) {
                    const _d = SSF.parse_date_code(p.v)
                    if (_d) {
                        p.t = 'd'
                        p.v = new Date(Date.UTC(_d.y, _d.m - 1, _d.d, _d.H, _d.M, _d.S, _d.u))
                    }
                }
                if (dense) {
                    const _r = decode_cell(tag.r)
                    if (!s[_r.r]) {
                        s[_r.r] = []
                    }
                    s[_r.r][_r.c] = p
                } else {
                    s[tag.r] = p
                }
            }
        }

        if (rows.length > 0) {
            s['!rows'] = rows
        }
    }
}()

function write_ws_xml_data(ws /*:Worksheet*/, opts, idx /*:number*/, wb /*:Workbook*/, rels) /*:string*/ {
    const o = []
    let r = []
    const range = safe_decode_range(ws['!ref'])
    let cell
    let ref
    let rr = ''
    const cols = []
    let R = 0
    let C = 0
    const rows = ws['!rows']
    const dense = Array.isArray(ws)
    for (C = range.s.c; C <= range.e.c; ++C) {
        cols[C] = encode_col(C)
    }
    for (R = range.s.r; R <= range.e.r; ++R) {
        r = []
        rr = encode_row(R)
        for (C = range.s.c; C <= range.e.c; ++C) {
            ref = cols[C] + rr
            const _cell = dense ? (ws[R] || [])[C] : ws[ref]
            if (_cell === undefined) {
                continue
            }
            if ((cell = write_ws_xml_cell(_cell, ref, ws, opts, idx, wb)) != null) {
                r.push(cell)
            }
        }
        if (r.length > 0) {
            const params = { r: rr }
            /*:any*/
            if (rows && rows[R]) {
                const row = rows[R]
                if (row.hidden) {
                    params.hidden = 1
                }
                let height = -1
                if (row.hpx) {
                    height = px2pt(row.hpx)
                } else if (row.hpt) {
                    height = row.hpt
                }
                if (height > -1) {
                    params.ht = height
                    params.customHeight = 1
                }
            }
            o[o.length] = writextag('row', r.join(''), params)
        }
    }
    return o.join('')
}

const WS_XML_ROOT = writextag('worksheet', null, {
    'xmlns': XMLNS.main[0],
    'xmlns:r': XMLNS.r,
})

export function write_ws_xml(idx /*:number*/, opts, wb /*:Workbook*/, rels) /*:string*/ {
    const o = [XML_HEADER, WS_XML_ROOT]
    const s = wb.SheetNames[idx]
    let sidx = 0
    let rdata = ''
    let ws = wb.Sheets[s]
    if (ws == null) {
        ws = {}
    }
    let ref = ws['!ref']
    if (ref == null) {
        ref = 'A1'
    }
    if (!rels) {
        rels = {}
    }
    ws['!comments'] = []
    ws['!drawing'] = []

    o[o.length] = writextag('sheetPr', null, { 'codeName': escapexml(wb.SheetNames[idx]) })

    o[o.length] = writextag('dimension', null, { 'ref': ref })

    o[o.length] = write_ws_xml_sheetviews(ws, opts, idx, wb)

    /* TODO: store in WB, process styles */
    if (opts.sheetFormat) {
        o[o.length] = writextag('sheetFormatPr', null, {
            defaultRowHeight: opts.sheetFormat.defaultRowHeight || '16',
            baseColWidth: opts.sheetFormat.baseColWidth || '10',
        })
    }

    if (ws['!cols'] != null && ws['!cols'].length > 0) {
        o[o.length] = write_ws_xml_cols(ws, ws['!cols'])
    }

    o[sidx = o.length] = '<sheetData/>'
    ws['!links'] = []
    if (ws['!ref'] != null) {
        rdata = write_ws_xml_data(ws, opts, idx, wb, rels)
        if (rdata.length > 0) {
            o[o.length] = rdata
        }
    }
    if (o.length > sidx + 1) {
        o[o.length] = '</sheetData>'
        o[sidx] = o[sidx].replace('/>', '>')
    }

    /* sheetCalcPr */

    if (ws['!protect'] != null) {
        o[o.length] = write_ws_xml_protection(ws['!protect'])
    }

    /* protectedRanges */
    /* scenarios */

    if (ws['!autofilter'] != null) {
        o[o.length] = write_ws_xml_autofilter(ws['!autofilter'])
    }

    /* sortState */
    /* dataConsolidate */
    /* customSheetViews */

    if (ws['!merges'] != null && ws['!merges'].length > 0) {
        o[o.length] = write_ws_xml_merges(ws['!merges'])
    }

    /* phoneticPr */
    /* conditionalFormatting */
    /* dataValidations */

    let relc = -1

    let rel
    let rId = -1
    if (ws['!links'].length > 0) {
        o[o.length] = '<hyperlinks>'
        ws['!links'].forEach(function (l) {
            if (!l[1].Target) {
                return
            }
            rId = add_rels(rels, -1, escapexml(l[1].Target).replace(/#.*$/, ''), RELS.HLINK)
            rel = { 'ref': l[0], 'r:id': `rId${rId}` }
            /*:any*/
            if ((relc = l[1].Target.indexOf('#')) > -1) {
                rel.location = escapexml(l[1].Target.substr(relc + 1))
            }
            if (l[1].Tooltip) {
                rel.tooltip = escapexml(l[1].Tooltip)
            }
            o[o.length] = writextag('hyperlink', null, rel)
        })
        o[o.length] = '</hyperlinks>'
    }
    delete ws['!links']

    /* printOptions */
    if (ws['!margins'] != null) {
        o[o.length] = write_ws_xml_margins(ws['!margins'])
    }
    /* pageSetup */

    const hfidx = o.length
    o[o.length] = ''

    /* rowBreaks */
    /* colBreaks */
    /* customProperties */
    /* cellWatches */
    /* ignoredErrors */
    /* smartTags */

    if (ws['!drawing'].length > 0) {
        rId = add_rels(rels, -1, `../drawings/drawing${idx + 1}.xml`, RELS.DRAW)
        ws['!drawing'].rid = rId
        o[o.length] = writextag('drawing', null, { 'r:id': `rId${rId}` })
    } else {
        delete ws['!drawing']
    }

    if (ws['!comments'].length > 0) {
        rId = add_rels(rels, -1, `../drawings/vmlDrawing${idx + 1}.vml`, RELS.VML)
        o[o.length] = writextag('legacyDrawing', null, { 'r:id': `rId${rId}` })
        ws['!legacy'] = rId
    }

    /* drawingHF */
    /* picture */
    /* oleObjects */
    /* controls */
    /* webPublishItems */
    /* tableParts */
    /* extList */

    if (o.length > 2) {
        o[o.length] = '</worksheet>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}
