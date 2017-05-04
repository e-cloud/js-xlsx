import { reset_cp, setCurrentCodepage } from './02_codepage'
import { make_ssf, SSF } from './10_ssf'
import { keys } from './20_jsutils'
import { getzipdata, getzipfile, getzipstr, resolve_path, safegetzipfile } from './21_ziputils'
import { parse_ct } from './30_ctype'
import { get_rels_path, parse_rels, RELS } from './31_rels'
import { parse_core_props } from './33_coreprops'
import { parse_ext_props } from './34_extprops'
import { parse_cust_props } from './35_custprops'
import {
    parse_DataSpaceDefinition,
    parse_DataSpaceMap,
    parse_DataSpaceVersionInfo,
    parse_EncryptionInfo,
    parse_Primary
} from './44_offcrypto'
import { parse_drawing } from './54_drawing'
import { parse_comments } from './56_cmntcommon'
import { strs } from './66_wscommon'
import { parse_chart } from './69_chartxml'
import {
    parse_cc,
    parse_cs,
    parse_ds,
    parse_ms,
    parse_sst,
    parse_sty,
    parse_theme,
    parse_wb,
    parse_ws
} from './74_xmlbin'
import { parse_ods } from './83_ods'
import { fix_read_opts } from './84_defaults'

function get_sheet_type(n) {
    if (RELS.WS.includes(n)) {
        return 'sheet'
    }
    if (RELS.CS && n == RELS.CS) {
        return 'chart'
    }
    if (RELS.DS && n == RELS.DS) {
        return 'dialog'
    }
    if (RELS.MS && n == RELS.MS) {
        return 'macro'
    }
    if (!n || !n.length) {
        return 'sheet'
    }
    return n
}

function safe_parse_wbrels(wbrels, sheets) {
    if (!wbrels) {
        return 0
    }
    try {
        wbrels = sheets.map(function pwbr(w) {
            if (!w.id) {
                w.id = w.strRelID
            }
            return [w.name, wbrels['!id'][w.id].Target, get_sheet_type(wbrels['!id'][w.id].Type)]
        })
    } catch (e) {
        return null
    }
    return !wbrels || wbrels.length === 0 ? null : wbrels
}

function safe_parse_sheet(
    zip,
    path
    /*:string*/,
    relsPath
    /*:string*/,
    sheet,
    sheetRels,
    sheets,
    stype
    /*:string*/,
    opts,
    wb,
    themes,
    styles,
) {
    try {
        sheetRels[sheet] = parse_rels(getzipstr(zip, relsPath, true), path)
        const data = getzipdata(zip, path)
        switch (stype) {
            case 'sheet':
                sheets[sheet] = parse_ws(data, path, opts, sheetRels[sheet], wb, themes, styles)
                break
            case 'chart':
                let cs = parse_cs(data, path, opts, sheetRels[sheet], wb, themes, styles)
                sheets[sheet] = cs
                if (!cs || !cs['!chart']) {
                    break
                }
                const dfile = resolve_path(cs['!chart'].Target, path)
                const drelsp = get_rels_path(dfile)
                const draw = parse_drawing(getzipstr(zip, dfile, true), parse_rels(getzipstr(zip, drelsp, true), dfile))
                const chartp = resolve_path(draw, dfile)
                const crelsp = get_rels_path(chartp)
                cs = parse_chart(getzipstr(zip, chartp, true), chartp, opts, parse_rels(getzipstr(zip, crelsp, true), chartp), wb, cs)
                break
            case 'macro':
                sheets[sheet] = parse_ms(data, path, opts, sheetRels[sheet], wb, themes, styles)
                break
            case 'dialog':
                sheets[sheet] = parse_ds(data, path, opts, sheetRels[sheet], wb, themes, styles)
                break
        }
    } catch (e) {
        if (opts.WTF) {
            throw e
        }
    }
}

const nodirs = function nodirs(x /*:string*/) /*:boolean*/ {
    return x.slice(-1) != '/'
}

export function parse_zip(zip /*:ZIP*/, opts /*:?ParseOpts*/) /*:Workbook*/ {
    make_ssf(SSF)
    opts = opts || {}
    fix_read_opts(opts)
    reset_cp()

    /* OpenDocument Part 3 Section 2.2.1 OpenDocument Package */
    if (safegetzipfile(zip, 'META-INF/manifest.xml')) {
        return parse_ods(zip, opts)
    }
    /* UOC */
    if (safegetzipfile(zip, 'objectdata.xml')) {
        return parse_ods(zip, opts)
    }

    const entries = keys(zip.files).filter(nodirs).sort()
    const dir = parse_ct(getzipstr(zip, '[Content_Types].xml') /*:?any*/, opts)
    let xlsb = false
    let sheets
    let binname
    if (dir.workbooks.length === 0) {
        binname = 'xl/workbook.xml'
        if (getzipdata(zip, binname, true)) {
            dir.workbooks.push(binname)
        }
    }
    if (dir.workbooks.length === 0) {
        binname = 'xl/workbook.bin'
        if (!getzipfile(zip, binname, true)) {
            throw new Error('Could not find workbook')
        }
        dir.workbooks.push(binname)
        xlsb = true
    }
    if (dir.workbooks[0].slice(-3) == 'bin') {
        xlsb = true
    }
    if (xlsb) {
        setCurrentCodepage(1200)
    }

    let themes = {}
    /*:any*/
    let styles = {}
    /*:any*/
    if (!opts.bookSheets && !opts.bookProps) {
        strs = []
        if (dir.sst) {
            strs = parse_sst(getzipdata(zip, dir.sst.replace(/^\//, '')), dir.sst, opts)
        }

        if (opts.cellStyles && dir.themes.length) {
            themes = parse_theme(getzipstr(zip, dir.themes[0].replace(/^\//, ''), true) || '', dir.themes[0], opts)
        }

        if (dir.style) {
            styles = parse_sty(getzipdata(zip, dir.style.replace(/^\//, '')), dir.style, themes, opts)
        }
    }

    const wb = parse_wb(getzipdata(zip, dir.workbooks[0].replace(/^\//, '')), dir.workbooks[0], opts)

    let props = {}
    let propdata = ''

    if (dir.coreprops.length !== 0) {
        propdata = getzipstr(zip, dir.coreprops[0].replace(/^\//, ''), true)
        if (propdata) {
            props = parse_core_props(propdata)
        }
        if (dir.extprops.length !== 0) {
            propdata = getzipstr(zip, dir.extprops[0].replace(/^\//, ''), true)
            if (propdata) {
                parse_ext_props(propdata, props)
            }
        }
    }

    let custprops = {}
    if (!opts.bookSheets || opts.bookProps) {
        if (dir.custprops.length !== 0) {
            propdata = getzipstr(zip, dir.custprops[0].replace(/^\//, ''), true)
            if (propdata) {
                custprops = parse_cust_props(propdata, opts)
            }
        }
    }

    let out = {}
    /*:any*/
    if (opts.bookSheets || opts.bookProps) {
        if (wb.Sheets) {
            sheets = wb.Sheets.map(function pluck(x) {
                return x.name
            })
        } else if (props.Worksheets && props.SheetNames.length > 0) {
            sheets = props.SheetNames
        }
        if (opts.bookProps) {
            out.Props = props
            out.Custprops = custprops
        }
        if (opts.bookSheets && typeof sheets !== 'undefined') {
            out.SheetNames = sheets
        }
        if (opts.bookSheets ? out.SheetNames : opts.bookProps) {
            return out
        }
    }
    sheets = {}

    let deps = {}
    if (opts.bookDeps && dir.calcchain) {
        deps = parse_cc(getzipdata(zip, dir.calcchain.replace(/^\//, '')), dir.calcchain, opts)
    }

    let i = 0
    const sheetRels = {}
    /*:any*/
    let path
    let relsPath

    const wbsheets = wb.Sheets
    props.Worksheets = wbsheets.length
    props.SheetNames = []
    for (let j = 0; j != wbsheets.length; ++j) {
        props.SheetNames[j] = wbsheets[j].name
    }

    const wbext = xlsb ? 'bin' : 'xml'
    const wbrelsfile = `xl/_rels/workbook.${wbext}.rels`
    let wbrels = parse_rels(getzipstr(zip, wbrelsfile, true), wbrelsfile)
    if (wbrels) {
        wbrels = safe_parse_wbrels(wbrels, wb.Sheets)
    }
    /* Numbers iOS hack */
    const nmode = getzipdata(zip, 'xl/worksheets/sheet.xml', true) ? 1 : 0
    for (i = 0; i != props.Worksheets; ++i) {
        let stype = 'sheet'
        if (wbrels && wbrels[i]) {
            path = `xl/${wbrels[i][1].replace(/[\/]?xl\//, '')}`
            stype = wbrels[i][2]
        } else {
            path = `xl/worksheets/sheet${i + 1 - nmode}.${wbext}`
            path = path.replace(/sheet0\./, 'sheet.')
        }
        relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, '$1/_rels/$3.rels')
        safe_parse_sheet(zip, path, relsPath, props.SheetNames[i], sheetRels, sheets, stype, opts, wb, themes, styles)
    }

    if (dir.comments) {
        parse_comments(zip, dir.comments, sheets, sheetRels, opts)
    }

    out = {
        Directory: dir,
        Workbook: wb,
        Props: props,
        Custprops: custprops,
        Deps: deps,
        Sheets: sheets,
        SheetNames: props.SheetNames,
        Strings: strs,
        Styles: styles,
        Themes: themes,
        SSF: SSF.get_table(),
    }
    /*:any*/
    if (opts.bookFiles) {
        out.keys = entries
        out.files = zip.files
    }
    if (opts.bookVBA) {
        if (dir.vba.length > 0) {
            out.vbaraw = getzipdata(zip, dir.vba[0].replace(/^\//, ''), true)
        } else if (dir.defaults && dir.defaults.bin === 'application/vnd.ms-office.vbaProject') {
            out.vbaraw = getzipdata(zip, 'xl/vbaProject.bin', true)
        }
    }
    return out
}

/* references to [MS-OFFCRYPTO] */
export function parse_xlsxcfb(cfb, opts /*:?ParseOpts*/) /*:Workbook*/ {
    let f = 'Version'
    let data = cfb.find(f)
    if (!data) {
        throw new Error(`ECMA-376 Encrypted file missing ${f}`)
    }
    const version = parse_DataSpaceVersionInfo(data.content)

    /* 2.3.4.1 */
    f = 'DataSpaceMap'
    data = cfb.find(f)
    if (!data) {
        throw new Error(`ECMA-376 Encrypted file missing ${f}`)
    }
    const dsm = parse_DataSpaceMap(data.content)
    if (dsm.length != 1 || dsm[0].comps.length != 1 || dsm[0].comps[0].t != 0 || dsm[0].name != 'StrongEncryptionDataSpace' || dsm[0].comps[0].v != 'EncryptedPackage') {
        throw new Error(`ECMA-376 Encrypted file bad ${f}`)
    }

    f = 'StrongEncryptionDataSpace'
    data = cfb.find(f)
    if (!data) {
        throw new Error(`ECMA-376 Encrypted file missing ${f}`)
    }
    const seds = parse_DataSpaceDefinition(data.content)
    if (seds.length != 1 || seds[0] != 'StrongEncryptionTransform') {
        throw new Error(`ECMA-376 Encrypted file bad ${f}`)
    }

    /* 2.3.4.3 */
    f = '!Primary'
    data = cfb.find(f)
    if (!data) {
        throw new Error(`ECMA-376 Encrypted file missing ${f}`)
    }
    const hdr = parse_Primary(data.content)

    f = 'EncryptionInfo'
    data = cfb.find(f)
    if (!data) {
        throw new Error(`ECMA-376 Encrypted file missing ${f}`)
    }
    const einfo = parse_EncryptionInfo(data.content)

    throw new Error('File is password-protected')
}
