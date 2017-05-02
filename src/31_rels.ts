import { keys } from './20_jsutils'
import { resolve_path } from './21_ziputils'
/* 9.3 Relationships */
import { parsexmltag, tagregex, writextag, XML_HEADER, XMLNS } from './22_xmlutils'

export const RELS = {
    WB: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    SHEET: 'http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    HLINK: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    VML: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
    VBA: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
}
/*:any*/

/* 9.3.3 Representing Relationships */
export function get_rels_path(file /*:string*/) /*:string*/ {
    const n = file.lastIndexOf('/')
    return `${file.substr(0, n + 1)}_rels/${file.substr(n + 1)}.rels`
}

export function parse_rels(data /*:?string*/, currentFilePath /*:string*/) {
    if (!data) return data
    if (currentFilePath.charAt(0) !== '/') {
        currentFilePath = `/${currentFilePath}`
    }
    const rels = {}
    const hash = {};

    (data.match(tagregex) || []).forEach(function (x) {
        const y = parsexmltag(x)
        /* 9.3.2.2 OPC_Relationships */
        if (y[0] === '<Relationship') {
            const rel = {}
            rel.Type = y.Type
            rel.Target = y.Target
            rel.Id = y.Id
            rel.TargetMode = y.TargetMode
            const canonictarget = y.TargetMode === 'External' ? y.Target : resolve_path(y.Target, currentFilePath)
            rels[canonictarget] = rel
            hash[y.Id] = rel
        }
    })
    rels['!id'] = hash
    return rels
}

XMLNS.RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'

const RELS_ROOT = writextag('Relationships', null, {
    //'xmlns:ns0': XMLNS.RELS,
    'xmlns': XMLNS.RELS,
})

/* TODO */
export function write_rels(rels) /*:string*/ {
    const o = [XML_HEADER, RELS_ROOT]
    keys(rels['!id']).forEach(function (rid) {
        o[o.length] = writextag('Relationship', null, rels['!id'][rid])
    })
    if (o.length > 2) {
        o[o.length] = '</Relationships>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}

export function add_rels(rels, rId, f, type, relobj?) /*:number*/ {
    if (!relobj) relobj = {}
    if (!rels['!id']) rels['!id'] = {}
    if (rId < 0) {
        for (rId = 1; rels['!id'][`rId${rId}`]; ++rId) {
        }
    }
    relobj.Id = `rId${rId}`
    relobj.Type = type
    relobj.Target = f
    if (relobj.Type == RELS.HLINK) relobj.TargetMode = 'External'
    if (rels['!id'][relobj.Id]) {
        throw new Error(`Cannot rewrite rId ${rId}`)
    }
    rels['!id'][relobj.Id] = relobj
    rels[(`/${relobj.Target}`).replace('//', '/')] = relobj
    return rId
}
