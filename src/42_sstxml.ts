import { isval } from './20_jsutils'
import {
    escapehtml,
    escapexml,
    matchtag,
    parsexmltag,
    tagregex,
    unescapexml,
    utf8read,
    writextag,
    XML_HEADER,
    XMLNS
} from './22_xmlutils'
import { RELS } from './31_rels'

/* 18.4.1 charset to codepage mapping */
export const CS2CP = {
    /*::[*/0 /*::]*/: 1252, /* ANSI */
    /*::[*/1 /*::]*/: 65001, /* DEFAULT */
    /*::[*/2 /*::]*/: 65001, /* SYMBOL */
    /*::[*/77 /*::]*/: 10000, /* MAC */
    /*::[*/128 /*::]*/: 932, /* SHIFTJIS */
    /*::[*/129 /*::]*/: 949, /* HANGUL */
    /*::[*/130 /*::]*/: 1361, /* JOHAB */
    /*::[*/134 /*::]*/: 936, /* GB2312 */
    /*::[*/136 /*::]*/: 950, /* CHINESEBIG5 */
    /*::[*/161 /*::]*/: 1253, /* GREEK */
    /*::[*/162 /*::]*/: 1254, /* TURKISH */
    /*::[*/163 /*::]*/: 1258, /* VIETNAMESE */
    /*::[*/177 /*::]*/: 1255, /* HEBREW */
    /*::[*/178 /*::]*/: 1256, /* ARABIC */
    /*::[*/186 /*::]*/: 1257, /* BALTIC */
    /*::[*/204 /*::]*/: 1251, /* RUSSIAN */
    /*::[*/222 /*::]*/: 874, /* THAI */
    /*::[*/238 /*::]*/: 1250, /* EASTEUROPE */
    /*::[*/255 /*::]*/: 1252, /* OEM */
    /*::[*/69 /*::]*/: 6969 /* MISC */
}
/*:any*/

/* Parse a list of <r> tags */
const parse_rs = function parse_rs_factory() {
    const tregex = matchtag('t')
    const rpregex = matchtag('rPr')
    const rregex = /<(?:\w+:)?r>/g
    const rend = /<\/(?:\w+:)?r>/
    const nlregex = /\r\n/g
    /* 18.4.7 rPr CT_RPrElt */
    const parse_rpr = function parse_rpr(rpr, intro, outro) {
        const font = {}
        let cp = 65001
        let align = ''
        const m = rpr.match(tagregex)
        let i = 0
        if (m) {
            for (; i != m.length; ++i) {
                const y = parsexmltag(m[i])
                switch (y[0].replace(/\w*:/g, '')) {
                    /* 18.8.12 condense CT_BooleanProperty */
                    /* ** not required . */
                    case '<condense':
                        break
                    /* 18.8.17 extend CT_BooleanProperty */
                    /* ** not required . */
                    case '<extend':
                        break
                    /* 18.8.36 shadow CT_BooleanProperty */
                    /* ** not required . */
                    case '<shadow':
                        if (!y.val) {
                            break
                        }
                    /* falls through */
                    case '<shadow>':
                    case '<shadow/>':
                        font.shadow = 1
                        break
                    case '</shadow>':
                        break

                    /* 18.4.1 charset CT_IntProperty TODO */
                    case '<charset':
                        if (y.val == '1') {
                            break
                        }
                        cp = CS2CP[parseInt(y.val, 10)]
                        break

                    /* 18.4.2 outline CT_BooleanProperty TODO */
                    case '<outline':
                        if (!y.val) {
                            break
                        }
                    /* falls through */
                    case '<outline>':
                    case '<outline/>':
                        font.outline = 1
                        break
                    case '</outline>':
                        break

                    /* 18.4.5 rFont CT_FontName */
                    case '<rFont':
                        font.name = y.val
                        break

                    /* 18.4.11 sz CT_FontSize */
                    case '<sz':
                        font.sz = y.val
                        break

                    /* 18.4.10 strike CT_BooleanProperty */
                    case '<strike':
                        if (!y.val) {
                            break
                        }
                    /* falls through */
                    case '<strike>':
                    case '<strike/>':
                        font.strike = 1
                        break
                    case '</strike>':
                        break

                    /* 18.4.13 u CT_UnderlineProperty */
                    case '<u':
                        if (!y.val) {
                            break
                        }
                        switch (y.val) {
                            case 'double':
                                font.uval = 'double'
                                break
                            case 'singleAccounting':
                                font.uval = 'single-accounting'
                                break
                            case 'doubleAccounting':
                                font.uval = 'double-accounting'
                                break
                        }
                    /* falls through */
                    case '<u>':
                    case '<u/>':
                        font.u = 1
                        break
                    case '</u>':
                        break

                    /* 18.8.2 b */
                    case '<b':
                        if (y.val == '0') {
                            break
                        }
                    /* falls through */
                    case '<b>':
                    case '<b/>':
                        font.b = 1
                        break
                    case '</b>':
                        break

                    /* 18.8.26 i */
                    case '<i':
                        if (y.val == '0') {
                            break
                        }
                    /* falls through */
                    case '<i>':
                    case '<i/>':
                        font.i = 1
                        break
                    case '</i>':
                        break

                    /* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
                    case '<color':
                        if (y.rgb) {
                            font.color = y.rgb.substr(2, 6)
                        }
                        break

                    /* 18.8.18 family ST_FontFamily */
                    case '<family':
                        font.family = y.val
                        break

                    /* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
                    case '<vertAlign':
                        align = y.val
                        break

                    /* 18.8.35 scheme CT_FontScheme TODO */
                    case '<scheme':
                        break

                    default:
                        if (y[0].charCodeAt(1) !== 47) {
                            throw `Unrecognized rich format ${y[0]}`
                        }
                }
            }
        }
        const style = []

        if (font.u) {
            style.push('text-decoration: underline;')
        }
        if (font.uval) {
            style.push(`text-underline-style:${font.uval};`)
        }
        if (font.sz) {
            style.push(`font-size:${font.sz};`)
        }
        if (font.outline) {
            style.push('text-effect: outline;')
        }
        if (font.shadow) {
            style.push('text-shadow: auto;')
        }
        intro.push(`<span style="${style.join('')}">`)

        if (font.b) {
            intro.push('<b>')
            outro.push('</b>')
        }
        if (font.i) {
            intro.push('<i>')
            outro.push('</i>')
        }
        if (font.strike) {
            intro.push('<s>')
            outro.push('</s>')
        }

        if (align == 'superscript') {
            align = 'sup'
        } else if (align == 'subscript') {
            align = 'sub'
        }
        if (align != '') {
            intro.push(`<${align}>`)
            outro.push(`</${align}>`)
        }

        outro.push('</span>')
        return cp
    }

    /* 18.4.4 r CT_RElt */
    function parse_r(r) {
        const terms = [[], '', []]

        /* 18.4.12 t ST_Xstring */
        const t = r.match(tregex)

        let cp = 65001
        if (!isval(t) /*:: || !t*/) {
            return ''
        }
        terms[1] = t[1]

        const rpr = r.match(rpregex)
        if (isval(rpr) /*:: && rpr*/) {
            cp = parse_rpr(rpr[1], terms[0], terms[2])
        }

        return terms[0].join('') + terms[1].replace(nlregex, '<br/>') + terms[2].join('')
    }

    return function parse_rs(rs) {
        return rs.replace(rregex, '').split(rend).map(parse_r).join('')
    }
}()

/* 18.4.8 si CT_Rst */
const sitregex = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g

const sirregex = /<(?:\w+:)?r>/
const sirphregex = /<(?:\w+:)?rPh.*?>(.*?)<\/(?:\w+:)?rPh>/g
export function parse_si(x, opts?) {
    const html = opts ? opts.cellHTML : true
    const z = {}
    if (!x) {
        return null
    }
    let y
    /* 18.4.12 t ST_Xstring (Plaintext String) */
    // TODO: is whitespace actually valid here?
    if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
        z.t = utf8read(unescapexml(x.substr(x.indexOf('>') + 1).split(/<\/(?:\w+:)?t>/)[0]))
        z.r = utf8read(x)
        if (html) {
            z.h = escapehtml(z.t)
        }
    }
    /* 18.4.4 r CT_RElt (Rich Text Run) */
    else if (y = x.match(sirregex)) {
        z.r = utf8read(x)
        z.t = utf8read(unescapexml((x.replace(sirphregex, '').match(sitregex) || []).join('').replace(tagregex, '')))
        if (html) {
            z.h = parse_rs(z.r)
        }
    }
    /* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
    /* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
    return z
}

/* 18.4 Shared String Table */
const sstr0 = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/
const sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g
const sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/
export function parse_sst_xml(data /*:string*/, opts) /*:SST*/ {
    const s /*:SST*/ = []
    /*:any*/
    let ss = ''
    if (!data) {
        return s
    }
    /* 18.4.9 sst CT_Sst */
    let sst = data.match(sstr0)
    if (isval(sst) /*:: && sst*/) {
        ss = sst[2].replace(sstr1, '').split(sstr2)
        for (let i = 0; i != ss.length; ++i) {
            const o = parse_si(ss[i].trim(), opts)
            if (o != null) {
                s[s.length] = o
            }
        }
        sst = parsexmltag(sst[1])
        s.Count = sst.count
        s.Unique = sst.uniqueCount
    }
    return s
}

RELS.SST = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const straywsregex = /^\s|\s$|[\t\n\r]/

export function write_sst_xml(sst /*:SST*/, opts) /*:string*/ {
    if (!opts.bookSST) {
        return ''
    }
    const o = [XML_HEADER]
    o[o.length] = writextag('sst', null, {
        xmlns: XMLNS.main[0],
        count: sst.Count,
        uniqueCount: sst.Unique,
    })
    for (let i = 0; i != sst.length; ++i) {
        if (sst[i] == null) {
            continue
        }
        const s /*:XLString*/ = sst[i]
        let sitag = '<si>'
        if (s.r) {
            sitag += s.r
        } else {
            sitag += '<t'
            if (!s.t) {
                s.t = ''
            }
            if (s.t.match(straywsregex)) {
                sitag += ' xml:space="preserve"'
            }
            sitag += `>${escapexml(s.t)}</t>`
        }
        sitag += '</si>'
        o[o.length] = sitag
    }
    if (o.length > 2) {
        o[o.length] = '</sst>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}
