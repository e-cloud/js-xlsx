import { keys, parseDate } from './20_jsutils'
import { parsexmlbool, parsexmltag, unescapexml, write_vt, writextag, XML_HEADER, XMLNS } from './22_xmlutils'
import { RELS } from './31_rels'

/* 15.2.12.2 Custom File Properties Part */
XMLNS.CUST_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
RELS.CUST_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'

const custregex = /<[^>]+>[^<]*/g
export function parse_cust_props(data: string, opts) {
    const p = {}
    let name = ''
    const m = data.match(custregex)
    if (m) {
        for (let i = 0; i != m.length; ++i) {
            const x = m[i]
            const y = parsexmltag(x)
            switch (y[0]) {
                case '<?xml':
                    break
                case '<Properties':
                    break
                case '<property':
                    name = y.name
                    break
                case '</property>':
                    name = null
                    break
                default:
                    if (x.indexOf('<vt:') === 0) {
                        const toks = x.split('>')
                        const type = toks[0].substring(4)
                        const text = toks[1]
                        /* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
                        switch (type) {
                            case 'lpstr':
                            case 'bstr':
                            case 'lpwstr':
                                p[name] = unescapexml(text)
                                break
                            case 'bool':
                                p[name] = parsexmlbool(text, '<vt:bool>')
                                break
                            case 'i1':
                            case 'i2':
                            case 'i4':
                            case 'i8':
                            case 'int':
                            case 'uint':
                                p[name] = parseInt(text, 10)
                                break
                            case 'r4':
                            case 'r8':
                            case 'decimal':
                                p[name] = parseFloat(text)
                                break
                            case 'filetime':
                            case 'date':
                                p[name] = parseDate(text)
                                break
                            case 'cy':
                            case 'error':
                                p[name] = unescapexml(text)
                                break
                            default:
                                if (opts.WTF && typeof console !== 'undefined') {
                                    console.warn('Unexpected', x, type, toks)
                                }
                        }
                    } else if (x.substr(0, 2) === '</') {
                    } else if (opts.WTF) {
                        throw new Error(x)
                    }
            }
        }
    }
    return p
}

const CUST_PROPS_XML_ROOT = writextag('Properties', null, {
    'xmlns': XMLNS.CUST_PROPS,
    'xmlns:vt': XMLNS.vt,
})

export function write_cust_props(cp, opts): string {
    const o = [XML_HEADER, CUST_PROPS_XML_ROOT]
    if (!cp) {
        return o.join('')
    }
    let pid = 1
    keys(cp).forEach(function custprop(k) {
        ++pid
        // $FlowIgnore
        o[o.length] = writextag('property', write_vt(cp[k]), {
            'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
            'pid': pid,
            'name': k,
        })
    })
    if (o.length > 2) {
        o[o.length] = '</Properties>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}
