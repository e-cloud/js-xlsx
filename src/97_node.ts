import { Readable } from 'stream'
import { decode_range, encode_col, safe_decode_range } from './27_csfutils'
import { HTML_ } from './79_html'
import { make_csv_row } from './90_utils'

const write_csv_stream = function (sheet /*:Worksheet*/, opts /*:?Sheet2CSVOpts*/) {
    const stream = new Readable()
    const out = ''
    const o = opts == null ? {} : opts
    if (sheet == null || sheet['!ref'] == null) {
        stream.push(null)
        return stream
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
    let R = r.s.r
    stream._read = function () {
        if (R > r.e.r) return stream.push(null)
        while (R <= r.e.r) {
            row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o)
            if (row == null) {
                ++R
                continue
            }
            if (o.strip) {
                row = row.replace(endregex, '')
            }
            stream.push(row + RS)
            ++R
            break
        }
    }
    return stream
}

const HTML_BEGIN = '<html><body><table>'
const HTML_END = '</table></body></html>'

const write_html_stream = function (sheet /*:Worksheet*/, opts) {
    const stream = new Readable()
    const o /*:Array<string>*/ = []
    const r = decode_range(sheet['!ref'])
    /*:Cell*/
    let cell

    o.dense = Array.isArray(sheet)
    stream.push(HTML_BEGIN)

    let R = r.s.r
    let end = false
    stream._read = function () {
        if (R > r.e.r) {
            if (!end) {
                end = true
                stream.push(HTML_END)
            }
            return stream.push(null)
        }
        while (R <= r.e.r) {
            stream.push(HTML_._row(sheet, r, R, o))
            ++R
            break
        }
    }
    return stream
}

export const stream = {
    to_html: write_html_stream,
    to_csv: write_csv_stream,
}


