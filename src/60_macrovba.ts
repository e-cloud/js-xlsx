import { RELS } from './31_rels'

RELS.DS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet'
RELS.MS = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet'

/* macro and dialog sheet stubs */
export function parse_ds_bin() {
    return { '!type': 'dialog' }
}

export function parse_ds_xml() {
    return { '!type': 'dialog' }
}

export function parse_ms_bin() {
    return { '!type': 'macro' }
}

export function parse_ms_xml() {
    return { '!type': 'macro' }
}
