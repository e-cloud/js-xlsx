import { DENSE } from './03_consts'
import { escapexml, parsexmltag, unescapexml, writextag } from './22_xmlutils'
import { decode_range, encode_cell, encode_range, format_cell, sheet_to_workbook } from './27_csfutils'

/* TODO: in browser attach to DOM; in node use an html parser */
export const HTML_ = function () {
    function html_to_sheet(str /*:string*/, _opts) /*:Workbook*/ {
        const opts = _opts || {}
        if (DENSE != null && opts.dense == null) opts.dense = DENSE
        const ws /*:Worksheet*/ = opts.dense ? [] /*:any*/ : {}
        /*:any*/
        let i = str.indexOf('<table')
        let j = str.indexOf('</table')
        if (i == -1 || j == -1) throw new Error('Invalid HTML: missing <table> / </table> pair')
        const rows = str.slice(i, j).split(/(:?<tr[^>]*>)/)
        let R = -1
        let C = 0
        let RS = 0
        let CS = 0
        const range = {s: {r: 10000000, c: 10000000}, e: {r: 0, c: 0}}
        const merges = []
        const midx = 0
        for (i = 0; i < rows.length; ++i) {
            const row = rows[i].trim()
            if (row.substr(0, 3) == '<tr') {
                ++R
                C = 0
                continue
            }
            if (row.substr(0, 3) != '<td') continue
            const cells = row.split('</td>')
            for (j = 0; j < cells.length; ++j) {
                const cell = cells[j].trim()
                if (cell.substr(0, 3) != '<td') continue
                let m = cell
                let cc = 0
                /* TODO: parse styles etc */
                while (m.charAt(0) == '<' && (cc = m.indexOf('>')) > -1) m = m.slice(cc + 1)
                while (m.includes('>')) m = m.slice(0, m.lastIndexOf('<'))
                const tag = parsexmltag(cell.slice(0, cell.indexOf('>')))
                CS = tag.colspan ? +tag.colspan : 1
                if ((RS = +tag.rowspan) > 0 || CS > 1) {
                    merges.push({
                        s: {r: R, c: C},
                        e: {r: R + (RS || 1) - 1, c: C + CS - 1},
                    })
                }
                /* TODO: generate stub cells */
                if (!m.length) {
                    C += CS
                    continue
                }
                m = unescapexml(m).replace(/[\r\n]/g, '')
                if (range.s.r > R) range.s.r = R
                if (range.e.r < R) range.e.r = R
                if (range.s.c > C) range.s.c = C
                if (range.e.c < C) range.e.c = C
                if (opts.dense) {
                    if (!ws[R]) ws[R] = []
                    if (Number(m) == Number(m)) {
                        ws[R][C] = {t: 'n', v: +m}
                    } else {
                        ws[R][C] = {t: 's', v: m}
                    }
                } else {
                    const coord /*:string*/ = encode_cell({r: R, c: C})
                    /* TODO: value parsing */
                    if (Number(m) == Number(m)) {
                        ws[coord] = {t: 'n', v: +m}
                    } else {
                        ws[coord] = {t: 's', v: m}
                    }
                }
                C += CS
            }
        }
        ws['!ref'] = encode_range(range)
        return ws
    }

    function html_to_book(str /*:string*/, opts) /*:Workbook*/ {
        return sheet_to_workbook(html_to_sheet(str, opts), opts)
    }

    function make_html_row(ws /*:Worksheet*/, r /*:Range*/, R /*:number*/, o) /*:string*/ {
        const M = ws['!merges'] || []
        const oo = []
        for (let C = r.s.c; C <= r.e.c; ++C) {
            let RS = 0
            let CS = 0
            for (let j = 0; j < M.length; ++j) {
                if (M[j].s.r > R || M[j].s.c > C) continue
                if (M[j].e.r < R || M[j].e.c < C) continue
                if (M[j].s.r < R || M[j].s.c < C) {
                    RS = -1
                    break
                }
                RS = M[j].e.r - M[j].s.r + 1
                CS = M[j].e.c - M[j].s.c + 1
                break
            }
            if (RS < 0) continue
            const coord = encode_cell({r: R, c: C})
            const cell = o.dense ? (ws[R] || [])[C] : ws[coord]
            if (!cell || cell.v == null) {
                oo.push('<td></td>')
                continue
            }
            /* TODO: html entities */
            const w = cell.h || escapexml(cell.w || (format_cell(cell), cell.w) || '')
            const sp = {}
            if (RS > 1) sp.rowspan = RS
            if (CS > 1) sp.colspan = CS
            oo.push(writextag('td', w, sp))
        }
        return `<tr>${oo.join('')}</tr>`
    }

    function sheet_to_html(ws /*:Worksheet*/, opts) /*:string*/ {
        const o /*:Array<string>*/ = []
        const r = decode_range(ws['!ref'])
        o.dense = Array.isArray(ws)
        for (let R = r.s.r; R <= r.e.r; ++R) o.push(make_html_row(ws, r, R, o));
        return `<html><body><table>${o.join('')}</table></body></html>`
    }

    return {
        to_workbook: html_to_book,
        to_sheet: html_to_sheet,
        _row: make_html_row,
        from_sheet: sheet_to_html,
    }
}()

export function parse_dom_table(table /*:HTMLElement*/, _opts /*:?any*/) /*:Worksheet*/ {
    const opts = _opts || {}
    if (DENSE != null) opts.dense = DENSE
    const ws /*:Worksheet*/ = opts.dense ? [] /*:any*/ : {}
    /*:any*/
    const rows = table.getElementsByTagName('tr')
    const range = {s: {r: 0, c: 0}, e: {r: rows.length - 1, c: 0}}
    const merges = []
    let midx = 0
    let R = 0
    let _C = 0
    let C = 0
    let RS = 0
    let CS = 0
    for (; R < rows.length; ++R) {
        const row = rows[R]
        const elts = row.children
        for (_C = C = 0; _C < elts.length; ++_C) {
            const elt = elts[_C]
            const v = elts[_C].innerText
            for (midx = 0; midx < merges.length; ++midx) {
                const m = merges[midx]
                if (m.s.c == C && m.s.r <= R && R <= m.e.r) {
                    C = m.e.c + 1
                    midx = -1
                }
            }
            /* TODO: figure out how to extract nonstandard mso- style */
            CS = +elt.getAttribute('colspan') || 1
            if ((RS = +elt.getAttribute('rowspan')) > 0 || CS > 1) {
                merges.push({
                    s: {r: R, c: C},
                    e: {r: R + (RS || 1) - 1, c: C + CS - 1},
                })
            }
            let o = {t: 's', v}
            if (v != null && v.length && !isNaN(Number(v))) {
                o = {t: 'n', v: Number(v)}
            }
            if (opts.dense) {
                if (!ws[R]) ws[R] = []
                ws[R][C] = o
            } else {
                ws[encode_cell({c: C, r: R})] = o
            }
            if (range.e.c < C) range.e.c = C
            C += CS
        }
    }
    ws['!merges'] = merges
    ws['!ref'] = encode_range(range)
    return ws
}

export function table_to_book(table /*:HTMLElement*/, opts /*:?any*/) /*:Workbook*/ {
    return sheet_to_workbook(parse_dom_table(table, opts), opts)
}
