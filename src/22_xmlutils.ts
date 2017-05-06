import { has_buf } from './05_buf'
import { evert, isval, keys } from './20_jsutils'

const attregexg = /([^"\s?>\/]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g
export const tagregex = /<[^>]*>/g
export const nsregex = /<\w*:/
const nsregex2 = /<(\/?)\w+:/

export function parsexmltag(tag: string, skip_root ?: boolean) {
    const z = {}
    let eq = 0
    let c = 0
    for (; eq !== tag.length; ++eq) {
        if ((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) {
            break
        }
    }
    if (!skip_root) {
        z[0] = tag.substr(0, eq)
    }
    if (eq === tag.length) {
        return z
    }
    const m = tag.match(attregexg)
    let j = 0
    let v = ''
    let i = 0
    let q = ''
    let cc = ''
    let quot = 1
    if (m) {
        for (i = 0; i != m.length; ++i) {
            cc = m[i]
            for (c = 0; c != cc.length; ++c) {
                if (cc.charCodeAt(c) === 61) {
                    break
                }
            }
            q = cc.substr(0, c)
            quot = (eq = cc.charCodeAt(c + 1)) == 34 || eq == 39 ? 1 : 0
            v = cc.substring(c + 1 + quot, cc.length - quot)
            for (j = 0; j != q.length; ++j) {
                if (q.charCodeAt(j) === 58) {
                    break
                }
            }
            if (j === q.length) {
                if (q.indexOf('_') > 0) {
                    q = q.substr(0, q.indexOf('_'))
                } // from ods
                z[q] = v
            } else {
                const k = (j === 5 && q.substr(0, 5) === 'xmlns' ? 'xmlns' : '') + q.substr(j + 1)
                if (z[k] && q.substr(j - 3, 3) == 'ext') {
                    continue
                } // from ods
                z[k] = v
            }
        }
    }
    return z
}
export function strip_ns(x: string): string {
    return x.replace(nsregex2, '<$1')
}

const encodings = {
    '&quot;': '"',
    '&apos;': '\'',
    '&gt;': '>',
    '&lt;': '<',
    '&amp;': '&',
}
const rencoding = evert(encodings)
const rencstr = '&<>\'"'.split('')

// TODO: CP remap (need to read file version to determine OS)
export const unescapexml: StringConv = function () {
    /* 22.4.2.4 bstr (Basic String) */
    const encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/g

    const coderegex = /_x([\da-fA-F]{4})_/g
    return function unescapexml(text: string): string {
        const s = `${text}`
        return s.replace(encregex, function ($$, $1) {
            return encodings[$$] || String.fromCharCode(parseInt($1, $$.includes('x') ? 16 : 10)) || $$
        }).replace(coderegex, function (m, c) {
            return String.fromCharCode(parseInt(c, 16))
        })
    }
}()

const decregex = /[&<>'"]/g
const charegex = /[\u0000-\u0008\u000b-\u001f]/g
export function escapexml(text: string, xml?: boolean): string {
    const s = `${text}`
    return s.replace(decregex, function (y) {
        return rencoding[y]
    }).replace(charegex, function (s) {
        return `_x${('000' + s.charCodeAt(0).toString(16)).slice(-4)}_`
    })
}
export function escapexmltag(text: string): string {
    return escapexml(text).replace(/ /g, '_x0020_')
}

const htmlcharegex = /[\u0000-\u001f]/g
export function escapehtml(text) {
    const s = `${text}`
    return s.replace(decregex, function (y) {
        return rencoding[y]
    }).replace(htmlcharegex, function (s) {
        return `&#x${('000' + s.charCodeAt(0).toString(16)).slice(-4)};`
    })
}

/* TODO: handle codepages */
export const xlml_fixstr: StringConv = function () {
    const entregex = /&#(\d+);/g

    function entrepl($$: string, $1: string): string {
        return String.fromCharCode(parseInt($1, 10))
    }

    return function xlml_fixstr(str: string): string {
        return str.replace(entregex, entrepl)
    }
}()
export const xlml_unfixstr: StringConv = function () {
    return function xlml_unfixstr(str: string): string {
        return str.replace(/(\r\n|[\r\n])/g, '&#10;')
    }
}()

export function parsexmlbool(value: any, tag?: string): boolean {
    switch (value) {
        case '1':
        case 'true':
        case 'TRUE':
            return true
        /* case '0': case 'false': case 'FALSE':*/
        default:
            return false
    }
}

export let utf8read: StringConv = function utf8reada(orig) {
    let out = ''
    let i = 0
    let c = 0
    let d = 0
    let e = 0
    let f = 0
    let w = 0
    while (i < orig.length) {
        c = orig.charCodeAt(i++)
        if (c < 128) {
            out += String.fromCharCode(c)
            continue
        }
        d = orig.charCodeAt(i++)
        if (c > 191 && c < 224) {
            out += String.fromCharCode((c & 31) << 6 | d & 63)
            continue
        }
        e = orig.charCodeAt(i++)
        if (c < 240) {
            out += String.fromCharCode((c & 15) << 12 | (d & 63) << 6 | e & 63)
            continue
        }
        f = orig.charCodeAt(i++)
        w = ((c & 7) << 18 | (d & 63) << 12 | (e & 63) << 6 | f & 63) - 65536
        out += String.fromCharCode(0xD800 + (w >>> 10 & 1023))
        out += String.fromCharCode(0xDC00 + (w & 1023))
    }
    return out
}

if (has_buf) {
    const utf8readb = function utf8readb(data) {
        const out = new Buffer(2 * data.length)
        let w
        let i
        let j = 1
        let k = 0
        let ww = 0
        let c
        for (i = 0; i < data.length; i += j) {
            j = 1
            if ((c = data.charCodeAt(i)) < 128) {
                w = c
            } else if (c < 224) {
                w = (c & 31) * 64 + (data.charCodeAt(i + 1) & 63)
                j = 2
            } else if (c < 240) {
                w = (c & 15) * 4096 + (data.charCodeAt(i + 1) & 63) * 64 + (data.charCodeAt(i + 2) & 63)
                j = 3
            } else {
                j = 4
                w = (c & 7) * 262144 + (data.charCodeAt(i + 1) & 63) * 4096 + (data.charCodeAt(i + 2) & 63) * 64 + (data.charCodeAt(i + 3) & 63)
                w -= 65536
                ww = 0xD800 + (w >>> 10 & 1023)
                w = 0xDC00 + (w & 1023)
            }
            if (ww !== 0) {
                out[k++] = ww & 255
                out[k++] = ww >>> 8
                ww = 0
            }
            out[k++] = w % 256
            out[k++] = w >>> 8
        }
        //out.length = k
        return out.toString('ucs2')
    }
    const corpus = 'foo bar baz\xE2\x98\x83\xF0\x9F\x8D\xA3'
    if (utf8read(corpus) == utf8readb(corpus)) {
        utf8read = utf8readb
    }
    // $FlowIgnore
    const utf8readc = function utf8readc(data) {
        return new Buffer(data, 'binary').toString('utf8')
    }
    if (utf8read(corpus) == utf8readc(corpus)) {
        utf8read = utf8readc
    }
}

// matches <foo>...</foo> extracts content
export const matchtag = function () {
    const mtcache: { [k: string]: RegExp } = {}
    return function matchtag(f, g?: string): RegExp {
        const t = `${f}|${g || ''}`
        if (mtcache[t]) {
            return mtcache[t]
        }
        return mtcache[t] = new RegExp(`<(?:\\w+:)?${f}(?: xml:space="preserve")?(?:[^>]*)>([^\u2603]*)</(?:\\w+:)?${f}>`, g || '')
    }
}()

const vtregex = function () {
    const vt_cache = {}
    return function vt_regex(bt) {
        if (vt_cache[bt] !== undefined) {
            return vt_cache[bt]
        }
        return vt_cache[bt] = new RegExp(`<(?:vt:)?${bt}>(.*?)</(?:vt:)?${bt}>`, 'g')
    }
}()

const vtvregex = /<\/?(?:vt:)?variant>/g
const vtmregex = /<(?:vt:)([^>]*)>(.*)</

export function parseVector(data) {
    const h = parsexmltag(data)
    const matches = data.match(vtregex(h.baseType)) || []
    if (matches.length != h.size) {
        throw new Error(`unexpected vector length ${matches.length} != ${h.size}`)
    }
    const res = []
    matches.forEach(function (x) {
        const v = x.replace(vtvregex, '').match(vtmregex)
        res.push({ v: utf8read(v[2]), t: v[1] })
    })
    return res
}

const wtregex = /(^\s|\s$|\n)/
export function writetag(f, g) {
    return `<${f}${g.match(wtregex) ? ' xml:space="preserve"' : ''}>${g}</${f}>`
}

export function wxt_helper(h): string {
    return keys(h).map(k => ` ${k}="${h[k]}"`).join('')
}

export function writextag(f, g, h?) {
    return `<${f}${isval(h) /*:: && h */ ? wxt_helper(h) : ''}${isval(g) /*:: && g */ ? (g.match(wtregex)
            ? ' xml:space="preserve"'
            : '') + `>${g}</${f}` : '/'}>`
}

export function write_w3cdtf(d: Date, t?: boolean): string {
    try {
        return d.toISOString().replace(/\.\d*/, '')
    } catch (e) {
        if (t) {
            throw e
        }
    }
    return ''
}

export function write_vt(s) {
    switch (typeof s) {
        case 'string':
            return writextag('vt:lpwstr', s)
        case 'number':
            return writextag((s | 0) == s ? 'vt:i4' : 'vt:r8', String(s))
        case 'boolean':
            return writextag('vt:bool', s ? 'true' : 'false')
    }
    if (s instanceof Date) {
        return writextag('vt:filetime', write_w3cdtf(s))
    }
    throw new Error(`Unable to serialize ${s}`)
}

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

export const XMLNS = {
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'mx': 'http://schemas.microsoft.com/office/mac/excel/2008/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'sjs': 'http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties',
    'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'xsd': 'http://www.w3.org/2001/XMLSchema',
    'main': [
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'http://purl.oclc.org/ooxml/spreadsheetml/main',
        'http://schemas.microsoft.com/office/excel/2006/main',
        'http://schemas.microsoft.com/office/excel/2006/2',
    ],
}

export const XLMLNS = {
    'o': 'urn:schemas-microsoft-com:office:office',
    'x': 'urn:schemas-microsoft-com:office:excel',
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
    'dt': 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882',
    'mv': 'http://macVmlSchemaUri',
    'v': 'urn:schemas-microsoft-com:vml',
    'html': 'http://www.w3.org/TR/REC-html40',
}
