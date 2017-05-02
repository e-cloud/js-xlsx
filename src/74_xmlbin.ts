import { parse_sst_xml, write_sst_xml } from './42_sstxml'
import { parse_sst_bin, write_sst_bin } from './43_sstbin'
import { parse_sty_xml, write_sty_xml } from './47_styxml'
import { parse_sty_bin, write_sty_bin } from './48_stybin'
import { parse_theme_xml } from './49_theme'
import { parse_cc_xml } from './52_ccxml'
import { parse_cc_bin } from './53_ccbin'
import { parse_comments_xml, write_comments_xml } from './57_cmntxml'
import { parse_comments_bin, write_comments_bin } from './58_cmntbin'
import { parse_ds_bin, parse_ds_xml, parse_ms_bin, parse_ms_xml } from './60_macrovba'
import { parse_ws_xml, write_ws_xml } from './67_wsxml'
import { parse_ws_bin, write_ws_bin } from './68_wsbin'
import { parse_cs_bin, parse_cs_xml, write_cs_bin, write_cs_xml } from './70_csheet'
import { parse_wb_xml, write_wb_xml } from './72_wbxml'
import { parse_wb_bin, write_wb_bin } from './73_wbbin'

export function parse_wb(data, name /*:string*/, opts) /*:WorkbookFile*/ {
    if (name.slice(-4) === '.bin') {
        return parse_wb_bin(data /*:any*/, opts)
    }
    return parse_wb_xml(data /*:any*/, opts)
}

export function parse_ws(data, name /*:string*/, opts, rels, wb, themes, styles) /*:Worksheet*/ {
    if (name.slice(-4) === '.bin') {
        return parse_ws_bin(data /*:any*/, opts, rels, wb, themes, styles)
    }
    return parse_ws_xml(data /*:any*/, opts, rels, wb, themes, styles)
}

export function parse_cs(data, name /*:string*/, opts, rels, wb, themes, styles) /*:Worksheet*/ {
    if (name.slice(-4) === '.bin') {
        return parse_cs_bin(data /*:any*/, opts, rels, wb, themes, styles)
    }
    return parse_cs_xml(data /*:any*/, opts, rels, wb, themes, styles)
}

export function parse_ms(data, name /*:string*/, opts, rels, wb, themes, styles) /*:Worksheet*/ {
    if (name.slice(-4) === '.bin') {
        return parse_ms_bin(data /*:any*/, opts, rels, wb, themes, styles)
    }
    return parse_ms_xml(data /*:any*/, opts, rels, wb, themes, styles)
}

export function parse_ds(data, name /*:string*/, opts, rels, wb, themes, styles) /*:Worksheet*/ {
    if (name.slice(-4) === '.bin') {
        return parse_ds_bin(data /*:any*/, opts, rels, wb, themes, styles)
    }
    return parse_ds_xml(data /*:any*/, opts, rels, wb, themes, styles)
}

export function parse_sty(data, name /*:string*/, themes, opts) {
    if (name.slice(-4) === '.bin') {
        return parse_sty_bin(data /*:any*/, themes, opts)
    }
    return parse_sty_xml(data /*:any*/, themes, opts)
}

export function parse_theme(data /*:string*/, name /*:string*/, opts) {
    return parse_theme_xml(data, opts)
}

export function parse_sst(data, name /*:string*/, opts) /*:SST*/ {
    if (name.slice(-4) === '.bin') {
        return parse_sst_bin(data /*:any*/, opts)
    }
    return parse_sst_xml(data /*:any*/, opts)
}

export function parse_cmnt(data, name /*:string*/, opts) {
    if (name.slice(-4) === '.bin') {
        return parse_comments_bin(data /*:any*/, opts)
    }
    return parse_comments_xml(data /*:any*/, opts)
}

export function parse_cc(data, name /*:string*/, opts) {
    if (name.slice(-4) === '.bin') {
        return parse_cc_bin(data /*:any*/, opts)
    }
    return parse_cc_xml(data /*:any*/, opts)
}

export function write_wb(wb, name /*:string*/, opts) {
    return (name.slice(-4) === '.bin' ? write_wb_bin : write_wb_xml)(wb, opts)
}

export function write_ws(data /*:number*/, name /*:string*/, opts, wb /*:Workbook*/, rels) {
    return (name.slice(-4) === '.bin' ? write_ws_bin : write_ws_xml)(data, opts, wb, rels)
}

export function write_cs(data /*:number*/, name /*:string*/, opts, wb /*:Workbook*/, rels) {
    return (name.slice(-4) === '.bin' ? write_cs_bin : write_cs_xml)(data, opts, wb, rels)
}

export function write_sty(data, name /*:string*/, opts) {
    return (name.slice(-4) === '.bin' ? write_sty_bin : write_sty_xml)(data, opts)
}

export function write_sst(data /*:SST*/, name /*:string*/, opts) {
    return (name.slice(-4) === '.bin' ? write_sst_bin : write_sst_xml)(data, opts)
}

export function write_cmnt(data /*:Array<any>*/, name /*:string*/, opts) {
    return (name.slice(-4) === '.bin' ? write_comments_bin : write_comments_xml)(data, opts)
}

/*
 function write_cc(data, name:string, opts) {
 return (name.slice(-4)===".bin" ? write_cc_bin : write_cc_xml)(data, opts);
 }
 */
