import { escapexml, matchtag, parseVector, writextag, XML_HEADER, XMLNS } from './22_xmlutils'
import { RELS } from './31_rels'

/* 15.2.12.3 Extended File Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
export const EXT_PROPS: Array<Array<string>> = [
    [
        'Application',
        'Application',
        'string',
    ],
    [
        'AppVersion',
        'AppVersion',
        'string',
    ],
    [
        'Company',
        'Company',
        'string',
    ],
    [
        'DocSecurity',
        'DocSecurity',
        'string',
    ],
    [
        'Manager',
        'Manager',
        'string',
    ],
    [
        'HyperlinksChanged',
        'HyperlinksChanged',
        'bool',
    ],
    [
        'SharedDoc',
        'SharedDoc',
        'bool',
    ],
    [
        'LinksUpToDate',
        'LinksUpToDate',
        'bool',
    ],
    [
        'ScaleCrop',
        'ScaleCrop',
        'bool',
    ],
    [
        'HeadingPairs',
        'HeadingPairs',
        'raw',
    ],
    [
        'TitlesOfParts',
        'TitlesOfParts',
        'raw',
    ],
]

XMLNS.EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
RELS.EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'

export function parse_ext_props(data, p) {
    const q = {}
    if (!p) {
        p = {}
    }

    EXT_PROPS.forEach(function (f) {
        switch (f[2]) {
            case 'string':
                p[f[1]] = (data.match(matchtag(f[0])) || [])[1]
                break
            case 'bool':
                p[f[1]] = (data.match(matchtag(f[0])) || [])[1] === 'true'
                break
            case 'raw':
                const cur = data.match(new RegExp(`<${f[0]}[^>]*>(.*)</${f[0]}>`))
                if (cur && cur.length > 0) {
                    q[f[1]] = cur[1]
                }
                break
        }
    })

    if (q.HeadingPairs && q.TitlesOfParts) {
        const v = parseVector(q.HeadingPairs)
        const parts = parseVector(q.TitlesOfParts).map(function (x) {
            return x.v
        })
        let idx = 0
        let len = 0
        for (let i = 0; i !== v.length; i += 2) {
            len = +v[i + 1].v
            switch (v[i].v) {
                case 'Worksheets':
                case '\u5DE5\u4F5C\u8868':
                case '\u041B\u0438\u0441\u0442\u044B':
                case '\u30EF\u30FC\u30AF\u30B7\u30FC\u30C8':
                case '\u05D2\u05DC\u05D9\u05D5\u05E0\u05D5\u05EA \u05E2\u05D1\u05D5\u05D3\u05D4':
                case 'Arbeitsbl\xE4tter':
                case '\xC7al\u0131\u015Fma Sayfalar\u0131':
                case 'Feuilles de calcul':
                case 'Fogli di lavoro':
                case 'Folhas de c\xE1lculo':
                case 'Planilhas':
                case 'Werkbladen':
                    p.Worksheets = len
                    p.SheetNames = parts.slice(idx, idx + len)
                    break

                case 'Named Ranges':
                case 'Benannte Bereiche':
                    p.NamedRanges = len
                    p.DefinedNames = parts.slice(idx, idx + len)
                    break

                case 'Charts':
                case 'Diagramme':
                    p.Chartsheets = len
                    p.ChartNames = parts.slice(idx, idx + len)
                    break
            }
            idx += len
        }
    }

    return p
}

const EXT_PROPS_XML_ROOT = writextag('Properties', null, {
    'xmlns': XMLNS.EXT_PROPS,
    'xmlns:vt': XMLNS.vt,
})

export function write_ext_props(cp, opts): string {
    const o = []
    const p = {}
    const W = writextag
    if (!cp) {
        cp = {}
    }
    cp.Application = 'SheetJS'
    o[o.length] = XML_HEADER
    o[o.length] = EXT_PROPS_XML_ROOT

    EXT_PROPS.forEach(function (f) {
        if (cp[f[1]] === undefined) {
            return
        }
        let v
        switch (f[2]) {
            case 'string':
                v = String(cp[f[1]])
                break
            case 'bool':
                v = cp[f[1]] ? 'true' : 'false'
                break
        }
        if (v !== undefined) {
            o[o.length] = W(f[0], v)
        }
    })

    /* TODO: HeadingPairs, TitlesOfParts */
    o[o.length] = W('HeadingPairs', W(
        'vt:vector',
        W('vt:variant', '<vt:lpstr>Worksheets</vt:lpstr>') + W('vt:variant', W('vt:i4', String(cp.Worksheets))),
        { size: 2, baseType: 'variant' },
    ))
    o[o.length] = W('TitlesOfParts', W(
        'vt:vector',
        cp.SheetNames.map(s => `<vt:lpstr>${escapexml(s)}</vt:lpstr>`).join(''),
        { size: cp.Worksheets, baseType: 'lpstr' },
    ))
    if (o.length > 2) {
        o[o.length] = '</Properties>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}
