/* Common Name -> XLML Name */
import { evert, keys } from './20_jsutils'
import { escapexmltag, writetag, writextag, XLMLNS } from './22_xmlutils'
import { CORE_PROPS } from './33_coreprops'
import { EXT_PROPS } from './34_extprops'

const XLMLDocPropsMap = {
    Category: 'Category',
    ContentStatus: 'ContentStatus', /* NOTE: missing from schema */
    Keywords: 'Keywords',
    LastAuthor: 'LastAuthor',
    LastPrinted: 'LastPrinted',
    RevNumber: 'Revision',
    Author: 'Author',
    Comments: 'Description',
    Identifier: 'Identifier', /* NOTE: missing from schema */
    Language: 'Language', /* NOTE: missing from schema */
    Subject: 'Subject',
    Title: 'Title',
    CreatedDate: 'Created',
    ModifiedDate: 'LastSaved',

    Application: 'AppName',
    AppVersion: 'Version',
    TotalTime: 'TotalTime',
    Manager: 'Manager',
    Company: 'Company',
}
const evert_XLMLDPM = evert(XLMLDocPropsMap)

export function xlml_set_prop(Props, tag /*:string*/, val) {
    tag = evert_XLMLDPM[tag] || tag
    Props[tag] = val
}

/* TODO: verify */
export function xlml_write_docprops(Props, opts) {
    const o = []
    CORE_PROPS.concat(EXT_PROPS).forEach(function (p) {
        if (Props[p[1]] == null) return
        let m = opts && opts.Props && opts.Props[p[1]] != null ? opts.Props[p[1]] : Props[p[1]]
        switch (p[2]) {
            case 'date':
                m = new Date(m).toISOString()
                break
        }
        if (typeof m == 'number') {
            m = String(m)
        } else if (m === true || m === false) {
            m = m ? '1' : '0'
        } else if (m instanceof Date) m = new Date(m).toISOString()
        o.push(writetag(XLMLDocPropsMap[p[1]] || p[1], m))
    })
    return writextag('DocumentProperties', o.join(''), {xmlns: XLMLNS.o})
}

export function xlml_write_custprops(Props, Custprops, opts) {
    const BLACKLIST = ['Worksheets', 'SheetNames']
    const T = 'CustomDocumentProperties'
    const o = []
    if (Props) {
        keys(Props).forEach(function (k) {
            /*:: if(!Props) return; */
            if (!Props.hasOwnProperty(k)) return
            for (let i = 0; i < CORE_PROPS.length; ++i) {
                if (k == CORE_PROPS[i][1]) return
            }
            for (let i = 0; i < EXT_PROPS.length; ++i) {
                if (k == EXT_PROPS[i][1]) return
            }
            for (let i = 0; i < BLACKLIST.length; ++i) {
                if (k == BLACKLIST[i]) return
            }

            let m = Props[k]
            let t = 'string'
            if (typeof m == 'number') {
                t = 'float'
                m = String(m)
            } else if (m === true || m === false) {
                t = 'boolean'
                m = m ? '1' : '0'
            } else {
                m = String(m)
            }
            o.push(writextag(escapexmltag(k), m, {'dt:dt': t}))
        })
    }
    if (Custprops) {
        keys(Custprops).forEach(function (k) {
            /*:: if(!Custprops) return; */
            if (!Custprops.hasOwnProperty(k)) return
            let m = Custprops[k]
            let t = 'string'
            if (typeof m == 'number') {
                t = 'float'
                m = String(m)
            } else if (m === true || m === false) {
                t = 'boolean'
                m = m ? '1' : '0'
            } else if (m instanceof Date) {
                t = 'dateTime.tz'
                m = m.toISOString()
            } else {
                m = String(m)
            }
            o.push(writextag(escapexmltag(k), m, {'dt:dt': t}))
        })
    }
    return `<${T} xmlns="${XLMLNS.o}">${o.join('')}</${T}>`
}
