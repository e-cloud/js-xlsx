import { chr0 } from './05_buf'
import { parsenoop } from './23_binutils'
import { parse_RkNumber, parse_Xnum } from './28_binstructs'
import { CountryEnum, XLSFillPattern } from './29_xlsenum'
import {
    parse_Bes,
    parse_ControlInfo,
    parse_Hyperlink,
    parse_LongRGB,
    parse_ShortXLUnicodeString,
    parse_XLUnicodeRichExtendedString,
    parse_XLUnicodeString,
    parse_XLUnicodeString2,
    parse_XLUnicodeStringNoCch,
    parsebool,
    parsenoop2,
    parseuint16,
    parseuint16a
} from './38_xlstypes'
import { parse_ArrayParsedFormula, parse_NameParsedFormula, parse_SharedParsedFormula } from './62_fxls'
/* --- MS-XLS --- */

/* 2.5.19 */
export function parse_XLSCell(blob, length?) /*:Cell*/ {
    const rw = blob.read_shift(2) // 0-indexed
    const col = blob.read_shift(2)
    const ixfe = blob.read_shift(2)
    return {r: rw, c: col, ixfe}
    /*:any*/
}

/* 2.5.134 */
export function parse_frtHeader(blob) {
    const rt = blob.read_shift(2)
    const flags = blob.read_shift(2) // TODO: parse these flags
    blob.l += 8
    return {type: rt, flags}
}

export function parse_OptXLUnicodeString(blob, length, opts) {
    return length === 0 ? '' : parse_XLUnicodeString2(blob, length, opts)
}

/* 2.5.158 */
const HIDEOBJENUM = ['SHOWALL', 'SHOWPLACEHOLDER', 'HIDEALL']
const parse_HideObjEnum = parseuint16

/* 2.5.344 */
export function parse_XTI(blob, length?) {
    const iSupBook = blob.read_shift(2)
    const itabFirst = blob.read_shift(2, 'i')
    const itabLast = blob.read_shift(2, 'i')
    return [iSupBook, itabFirst, itabLast]
}

/* 2.5.218 */
export function parse_RkRec(blob, length?) {
    const ixfe = blob.read_shift(2)
    const RK = parse_RkNumber(blob)
    return [ixfe, RK]
}

/* 2.5.1 */
function parse_AddinUdf(blob, length, opts) {
    blob.l += 4
    length -= 4
    let l = blob.l + length
    const udfName = parse_ShortXLUnicodeString(blob, length, opts)
    const cb = blob.read_shift(2)
    l -= blob.l
    if (cb !== l) {
        throw new Error(`Malformed AddinUdf: padding = ${l} != ${cb}`)
    }
    blob.l += cb
    return udfName
}

/* 2.5.209 TODO: Check sizes */
export function parse_Ref8U(blob, length) {
    const rwFirst = blob.read_shift(2)
    const rwLast = blob.read_shift(2)
    const colFirst = blob.read_shift(2)
    const colLast = blob.read_shift(2)
    return {s: {c: colFirst, r: rwFirst}, e: {c: colLast, r: rwLast}}
}

/* 2.5.211 */
export function parse_RefU(blob, length) {
    const rwFirst = blob.read_shift(2)
    const rwLast = blob.read_shift(2)
    const colFirst = blob.read_shift(1)
    const colLast = blob.read_shift(1)
    return {s: {c: colFirst, r: rwFirst}, e: {c: colLast, r: rwLast}}
}

/* 2.5.207 */
const parse_Ref = parse_RefU

/* 2.5.143 */
function parse_FtCmo(blob, length) {
    blob.l += 4
    const ot = blob.read_shift(2)
    const id = blob.read_shift(2)
    const flags = blob.read_shift(2)
    blob.l += 12
    return [id, ot, flags]
}

/* 2.5.149 */
function parse_FtNts(blob, length) {
    const out = {}
    blob.l += 4
    blob.l += 16 // GUID TODO
    out.fSharedNote = blob.read_shift(2)
    blob.l += 4
    return out
}

/* 2.5.142 */
function parse_FtCf(blob, length) {
    const out = {}
    blob.l += 4
    blob.cf = blob.read_shift(2)
    return out
}

/* 2.5.140 - 2.5.154 and friends */
const FtTab = {
    /*::[*/0x15 /*::]*/: parse_FtCmo,
    /*::[*/0x13 /*::]*/: parsenoop, /* FtLbsData */
    /*::[*/0x12 /*::]*/: function (blob, length) {
        blob.l += 12
    }, /* FtCblsData */
    /*::[*/0x11 /*::]*/: function (blob, length) {
        blob.l += 8
    }, /* FtRboData */
    /*::[*/0x10 /*::]*/: parsenoop, /* FtEdoData */
    /*::[*/0x0F /*::]*/: parsenoop, /* FtGboData */
    /*::[*/0x0D /*::]*/: parse_FtNts, /* FtNts */
    /*::[*/0x0C /*::]*/: function (blob, length) {
        blob.l += 24
    }, /* FtSbs */
    /*::[*/0x0B /*::]*/: function (blob, length) {
        blob.l += 10
    }, /* FtRbo */
    /*::[*/0x0A /*::]*/: function (blob, length) {
        blob.l += 16
    }, /* FtCbls */
    /*::[*/0x09 /*::]*/: parsenoop, /* FtPictFmla */
    /*::[*/0x08 /*::]*/: function (blob, length) {
        blob.l += 6
    }, /* FtPioGrbit */
    /*::[*/0x07 /*::]*/: parse_FtCf, /* FtCf */
    /*::[*/0x06 /*::]*/: function (blob, length) {
        blob.l += 6
    }, /* FtGmo */
    /*::[*/0x04 /*::]*/: parsenoop, /* FtMacro */
    /*::[*/0x00 /*::]*/: function (blob, length) {
        blob.l += 4
    } /* FtEnding */
}
function parse_FtArray(blob, length, ot) {
    const s = blob.l
    const fts = []
    while (blob.l < s + length) {
        const ft = blob.read_shift(2)
        blob.l -= 2
        try {
            fts.push(FtTab[ft](blob, s + length - blob.l))
        } catch (e) {
            blob.l = s + length
            return fts
        }
    }
    if (blob.l != s + length) {
        blob.l = s + length
    } //throw new Error("bad Object Ft-sequence");
    return fts
}

/* 2.5.129 */
const parse_FontIndex = parseuint16

/* --- 2.4 Records --- */

/* 2.4.21 */
export function parse_BOF(blob, length) {
    const o = {BIFFVer: 0, dt: 0}
    o.BIFFVer = blob.read_shift(2)
    length -= 2
    if (length >= 2) {
        o.dt = blob.read_shift(2)
        blob.l -= 2
    }
    switch (o.BIFFVer) {
        case 0x0600: /* BIFF8 */
        case 0x0500: /* BIFF5 */
        case 0x0002:
        case 0x0007:
            /* BIFF2 */
            break
        default:
            if (length > 6) {
                throw new Error(`Unexpected BIFF Ver ${o.BIFFVer}`)
            }
    }

    blob.read_shift(length)
    return o
}

/* 2.4.146 */
export function parse_InterfaceHdr(blob, length) {
    if (length === 0) return 0x04b0
    let q
    if ((q = blob.read_shift(2)) !== 0x04b0) {
    }
    return 0x04b0
}

/* 2.4.349 */
export function parse_WriteAccess(blob, length, opts) {
    if (opts.enc) {
        blob.l += length
        return ''
    }
    const l = blob.l
    // TODO: make sure XLUnicodeString doesnt overrun
    const UserName = parse_XLUnicodeString(blob, 0, opts)
    blob.read_shift(length + l - blob.l)
    return UserName
}

/* 2.4.28 */
export function parse_BoundSheet8(blob, length, opts) {
    const pos = blob.read_shift(4)
    const hidden = blob.read_shift(1) & 0x03
    let dt = blob.read_shift(1)
    switch (dt) {
        case 0:
            dt = 'Worksheet'
            break
        case 1:
            dt = 'Macrosheet'
            break
        case 2:
            dt = 'Chartsheet'
            break
        case 6:
            dt = 'VBAModule'
            break
    }
    let name = parse_ShortXLUnicodeString(blob, 0, opts)
    if (name.length === 0) name = 'Sheet1'
    return {pos, hs: hidden, dt, name}
}

/* 2.4.265 TODO */
export function parse_SST(blob, length) /*:SST*/ {
    const cnt = blob.read_shift(4)
    const ucnt = blob.read_shift(4)
    const strs /*:SST*/ = []
    /*:any*/
    for (let i = 0; i != ucnt; ++i) {
        strs.push(parse_XLUnicodeRichExtendedString(blob))
    }
    strs.Count = cnt
    strs.Unique = ucnt
    return strs
}

/* 2.4.107 */
export function parse_ExtSST(blob, length) {
    const extsst = {}
    extsst.dsst = blob.read_shift(2)
    blob.l += length - 2
    return extsst
}

/* 2.4.221 TODO: check BIFF2-4 */
export function parse_Row(blob, length) {
    const z = ({}/*:any*/)
    z.r = blob.read_shift(2)
    z.c = blob.read_shift(2)
    z.cnt = blob.read_shift(2) - z.c
    const miyRw = blob.read_shift(2)
    blob.l += 4 // reserved(2), unused(2)
    const flags = blob.read_shift(1) // various flags
    blob.l += 3 // reserved(8), ixfe(12), flags(4)
    if (flags & 0x20) z.hidden = true
    if (flags & 0x40) z.hpt = miyRw / 20
    return z
}

/* 2.4.125 */
export function parse_ForceFullCalculation(blob, length) {
    const header = parse_frtHeader(blob)
    if (header.type != 0x08A3) {
        throw new Error(`Invalid Future Record ${header.type}`)
    }
    const fullcalc = blob.read_shift(4)
    return fullcalc !== 0x0
}

export const parse_CompressPictures = parsenoop2
/* 2.4.55 Not interesting */

/* 2.4.215 rt */
export function parse_RecalcId(blob, length) {
    blob.read_shift(2)
    return blob.read_shift(4)
}

/* 2.4.87 */
export function parse_DefaultRowHeight(blob, length) {
    const f = blob.read_shift(2)
    const fl = {Unsynced: f & 1, DyZero: (f & 2) >> 1, ExAsc: (f & 4) >> 2, ExDsc: (f & 8) >> 3}
    /* char is misleading, miyRw and miyRwHidden overlap */
    const miyRw = blob.read_shift(2)
    return [fl, miyRw]
}

/* 2.4.345 TODO */
export function parse_Window1(blob, length) {
    const xWn = blob.read_shift(2)
    const yWn = blob.read_shift(2)
    const dxWn = blob.read_shift(2)
    const dyWn = blob.read_shift(2)
    const flags = blob.read_shift(2)
    const iTabCur = blob.read_shift(2)
    const iTabFirst = blob.read_shift(2)
    const ctabSel = blob.read_shift(2)
    const wTabRatio = blob.read_shift(2)
    return {
        Pos: [xWn, yWn], Dim: [dxWn, dyWn], Flags: flags, CurTab: iTabCur,
        FirstTab: iTabFirst, Selected: ctabSel, TabRatio: wTabRatio,
    }
}

/* 2.4.122 TODO */
export function parse_Font(blob, length, opts) {
    blob.l += 14
    const name = parse_ShortXLUnicodeString(blob, 0, opts)
    return name
}

/* 2.4.149 */
export function parse_LabelSst(blob, length?) {
    const cell = parse_XLSCell(blob)
    cell.isst = blob.read_shift(4)
    return cell
}

/* 2.4.148 */
export function parse_Label(blob, length, opts) {
    const target = blob.l + length
    const cell = parse_XLSCell(blob, 6)
    if (opts.biff == 2) blob.l++
    const str = parse_XLUnicodeString(blob, target - blob.l, opts)
    cell.val = str
    return cell
}

/* 2.4.126 Number Formats */
export function parse_Format(blob, length, opts) {
    const ifmt = blob.read_shift(2)
    const fmtstr = parse_XLUnicodeString2(blob, 0, opts)
    return [ifmt, fmtstr]
}
export const parse_BIFF2Format = parse_XLUnicodeString2

/* 2.4.90 */
export function parse_Dimensions(blob, length, opts) {
    const end = blob.l + length
    const w = opts.biff == 8 || !opts.biff ? 4 : 2
    const r = blob.read_shift(w)
    const R = blob.read_shift(w)
    const c = blob.read_shift(2)
    const C = blob.read_shift(2)
    blob.l = end
    return {s: {r, c}, e: {r: R, c: C}}
}

/* 2.4.220 */
export function parse_RK(blob, length) {
    const rw = blob.read_shift(2)
    const col = blob.read_shift(2)
    const rkrec = parse_RkRec(blob)
    return {r: rw, c: col, ixfe: rkrec[0], rknum: rkrec[1]}
}

/* 2.4.175 */
export function parse_MulRk(blob, length) {
    const target = blob.l + length - 2
    const rw = blob.read_shift(2)
    const col = blob.read_shift(2)
    const rkrecs = []
    while (blob.l < target) rkrecs.push(parse_RkRec(blob))
    if (blob.l !== target) {
        throw new Error('MulRK read error')
    }
    const lastcol = blob.read_shift(2)
    if (rkrecs.length != lastcol - col + 1) {
        throw new Error('MulRK length mismatch')
    }
    return {r: rw, c: col, C: lastcol, rkrec: rkrecs}
}
/* 2.4.174 */
export function parse_MulBlank(blob, length) {
    const target = blob.l + length - 2
    const rw = blob.read_shift(2)
    const col = blob.read_shift(2)
    const ixfes = []
    while (blob.l < target) ixfes.push(blob.read_shift(2))
    if (blob.l !== target) {
        throw new Error('MulBlank read error')
    }
    const lastcol = blob.read_shift(2)
    if (ixfes.length != lastcol - col + 1) {
        throw new Error('MulBlank length mismatch')
    }
    return {r: rw, c: col, C: lastcol, ixfe: ixfes}
}

/* 2.5.20 2.5.249 TODO: interpret values here */
function parse_CellStyleXF(blob, length, style, opts) {
    const o = {}
    const a = blob.read_shift(4)
    const b = blob.read_shift(4)
    const c = blob.read_shift(4)
    const d = blob.read_shift(2)
    o.patternType = XLSFillPattern[c >> 26]

    if (!opts.cellStyles) return o
    o.alc = a & 0x07
    o.fWrap = a >> 3 & 0x01
    o.alcV = a >> 4 & 0x07
    o.fJustLast = a >> 7 & 0x01
    o.trot = a >> 8 & 0xFF
    o.cIndent = a >> 16 & 0x0F
    o.fShrinkToFit = a >> 20 & 0x01
    o.iReadOrder = a >> 22 & 0x02
    o.fAtrNum = a >> 26 & 0x01
    o.fAtrFnt = a >> 27 & 0x01
    o.fAtrAlc = a >> 28 & 0x01
    o.fAtrBdr = a >> 29 & 0x01
    o.fAtrPat = a >> 30 & 0x01
    o.fAtrProt = a >> 31 & 0x01

    o.dgLeft = b & 0x0F
    o.dgRight = b >> 4 & 0x0F
    o.dgTop = b >> 8 & 0x0F
    o.dgBottom = b >> 12 & 0x0F
    o.icvLeft = b >> 16 & 0x7F
    o.icvRight = b >> 23 & 0x7F
    o.grbitDiag = b >> 30 & 0x03

    o.icvTop = c & 0x7F
    o.icvBottom = c >> 7 & 0x7F
    o.icvDiag = c >> 14 & 0x7F
    o.dgDiag = c >> 21 & 0x0F

    o.icvFore = d & 0x7F
    o.icvBack = d >> 7 & 0x7F
    o.fsxButton = d >> 14 & 0x01
    return o
}
export function parse_CellXF(blob, length, opts) {
    return parse_CellStyleXF(blob, length, 0, opts)
}
export function parse_StyleXF(blob, length, opts) {
    return parse_CellStyleXF(blob, length, 1, opts)
}

/* 2.4.353 TODO: actually do this right */
export function parse_XF(blob, length, opts) {
    const o = {}
    o.ifnt = blob.read_shift(2)
    o.ifmt = blob.read_shift(2)
    o.flags = blob.read_shift(2)
    o.fStyle = o.flags >> 2 & 0x01
    length -= 6
    o.data = parse_CellStyleXF(blob, length, o.fStyle, opts)
    return o
}

/* 2.4.134 */
export function parse_Guts(blob, length) {
    blob.l += 4
    const out = [blob.read_shift(2), blob.read_shift(2)]
    if (out[0] !== 0) out[0]--
    if (out[1] !== 0) out[1]--
    if (out[0] > 7 || out[1] > 7) {
        throw new Error(`Bad Gutters: ${out.join('|')}`)
    }
    return out
}

/* 2.4.24 */
export function parse_BoolErr(blob, length, opts) {
    const cell = parse_XLSCell(blob, 6)
    if (opts.biff == 2) ++blob.l
    const val = parse_Bes(blob, 2)
    cell.val = val
    cell.t = val === true || val === false ? 'b' : 'e'
    return cell
}

/* 2.4.180 Number */
export function parse_Number(blob, length) {
    const cell = parse_XLSCell(blob, 6)
    const xnum = parse_Xnum(blob, 8)
    cell.val = xnum
    return cell
}

const parse_XLHeaderFooter = parse_OptXLUnicodeString // TODO: parse 2.4.136

/* 2.4.271 */
export function parse_SupBook(blob, length, opts) {
    const end = blob.l + length
    const ctab = blob.read_shift(2)
    const cch = blob.read_shift(2)
    let virtPath
    if (cch >= 0x01 && cch <= 0xff) virtPath = parse_XLUnicodeStringNoCch(blob, cch)
    const rgst = blob.read_shift(end - blob.l)
    opts.sbcch = cch
    return [cch, ctab, virtPath, rgst]
}

/* 2.4.105 TODO */
export function parse_ExternName(blob, length, opts) {
    const flags = blob.read_shift(2)
    let body
    const o = {
        fBuiltIn: flags & 0x01,
        fWantAdvise: flags >>> 1 & 0x01,
        fWantPict: flags >>> 2 & 0x01,
        fOle: flags >>> 3 & 0x01,
        fOleLink: flags >>> 4 & 0x01,
        cf: flags >>> 5 & 0x3FF,
        fIcon: flags >>> 15 & 0x01,
    }
    /*:any*/
    if (opts.sbcch === 0x3A01) body = parse_AddinUdf(blob, length - 2, opts)
    //else throw new Error("unsupported SupBook cch: " + opts.sbcch);
    o.body = body || blob.read_shift(length - 2)
    if (typeof body === 'string') {
        o.Name = body
    }
    return o
}

/* 2.4.150 TODO */
export function parse_Lbl(blob, length, opts) {
    const target = blob.l + length
    const flags = blob.read_shift(2)
    const chKey = blob.read_shift(1)
    const cch = blob.read_shift(1)
    const cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2)
    let itab = 0
    if (!opts || opts.biff >= 5) {
        blob.l += 2
        itab = blob.read_shift(2)
        blob.l += 4
    }
    const name = parse_XLUnicodeStringNoCch(blob, cch, opts)
    let npflen = target - blob.l
    if (opts && opts.biff == 2) --npflen
    const rgce = target == blob.l || cce == 0 ? [] : parse_NameParsedFormula(blob, npflen, opts, cce)
    return {
        chKey,
        Name: name,
        itab,
        rgce,
    }
}

/* 2.4.106 TODO: verify supbook manipulation */
export function parse_ExternSheet(blob, length, opts) {
    if (opts.biff < 8) {
        return parse_ShortXLUnicodeString(blob, length, opts)
    }
    const o = []
    const target = blob.l + length
    let len = blob.read_shift(2)
    while (len-- !== 0) o.push(parse_XTI(blob, 6))
    // [iSupBook, itabFirst, itabLast];
    const oo = []
    return o
}

/* 2.4.176 TODO: check older biff */
export function parse_NameCmt(blob, length, opts) {
    if (opts.biff < 8) {
        blob.l += length
        return
    }
    const cchName = blob.read_shift(2)
    const cchComment = blob.read_shift(2)
    const name = parse_XLUnicodeStringNoCch(blob, cchName, opts)
    const comment = parse_XLUnicodeStringNoCch(blob, cchComment, opts)
    return [name, comment]
}

/* 2.4.260 */
export function parse_ShrFmla(blob, length, opts) {
    const ref = parse_RefU(blob, 6)
    blob.l++
    const cUse = blob.read_shift(1)
    length -= 8
    return [parse_SharedParsedFormula(blob, length, opts), cUse]
}

/* 2.4.4 TODO */
export function parse_Array(blob, length, opts) {
    const ref = parse_Ref(blob, 6)
    /* TODO: fAlwaysCalc */
    switch (opts.biff) {
        case 2:
            blob.l++
            length -= 7
            break
        case 3:
        case 4:
            blob.l += 2
            length -= 8
            break
        default:
            blob.l += 6
            length -= 12
    }
    return [ref, parse_ArrayParsedFormula(blob, length, opts, ref)]
}

/* 2.4.173 */
export function parse_MTRSettings(blob, length) {
    const fMTREnabled = blob.read_shift(4) !== 0x00
    const fUserSetThreadCount = blob.read_shift(4) !== 0x00
    const cUserThreadCount = blob.read_shift(4)
    return [fMTREnabled, fUserSetThreadCount, cUserThreadCount]
}

/* 2.5.186 TODO: BIFF5 */
export function parse_NoteSh(blob, length, opts) {
    if (opts.biff < 8) return
    const row = blob.read_shift(2)
    const col = blob.read_shift(2)
    const flags = blob.read_shift(2)
    const idObj = blob.read_shift(2)
    const stAuthor = parse_XLUnicodeString2(blob, 0, opts)
    if (opts.biff < 8) blob.read_shift(1)
    return [{r: row, c: col}, stAuthor, idObj, flags]
}

/* 2.4.179 */
export function parse_Note(blob, length, opts) {
    /* TODO: Support revisions */
    return parse_NoteSh(blob, length, opts)
}

/* 2.4.168 */
export function parse_MergeCells(blob, length) {
    const merges = []
    let cmcs = blob.read_shift(2)
    while (cmcs--) merges.push(parse_Ref8U(blob, length))
    return merges
}

/* 2.4.181 TODO: parse all the things! */
export function parse_Obj(blob, length) {
    const cmo = parse_FtCmo(blob, 22) // id, ot, flags
    const fts = parse_FtArray(blob, length - 22, cmo[1])
    return {cmo, ft: fts}
}

/* 2.4.329 TODO: parse properly */
export function parse_TxO(blob, length, opts) {
    const s = blob.l
    let texts = ''
    try {
        blob.l += 4
        const ot = (opts.lastobj || {cmo: [0, 0]}).cmo[1]
        let controlInfo
        if (![0, 5, 7, 11, 12, 14].includes(ot)) {
            blob.l += 6
        } else {
            controlInfo = parse_ControlInfo(blob, 6, opts)
        }
        const cchText = blob.read_shift(2)
        const cbRuns = blob.read_shift(2)
        const ifntEmpty = parse_FontIndex(blob, 2)
        const len = blob.read_shift(2)
        blob.l += len
        //var fmla = parse_ObjFmla(blob, s + length - blob.l);

        for (let i = 1; i < blob.lens.length - 1; ++i) {
            if (blob.l - s != blob.lens[i]) {
                throw new Error('TxO: bad continue record')
            }
            const hdr = blob[blob.l]
            const t = parse_XLUnicodeStringNoCch(blob, blob.lens[i + 1] - blob.lens[i] - 1)
            texts += t
            if (texts.length >= (hdr ? cchText : 2 * cchText)) break
        }
        if (texts.length !== cchText && texts.length !== cchText * 2) {
            throw new Error(`cchText: ${cchText} != ${texts.length}`)
        }

        blob.l = s + length
        /* 2.5.272 TxORuns */
        //	var rgTxoRuns = [];
        //	for(var j = 0; j != cbRuns/8-1; ++j) blob.l += 8;
        //	var cchText2 = blob.read_shift(2);
        //	if(cchText2 !== cchText) throw new Error("TxOLastRun mismatch: " + cchText2 + " " + cchText);
        //	blob.l += 6;
        //	if(s + length != blob.l) throw new Error("TxO " + (s + length) + ", at " + blob.l);
        return {t: texts}
    } catch (e) {
        blob.l = s + length
        return {t: texts}
    }
}

/* 2.4.140 */
export const parse_HLink = function (blob, length) {
    const ref = parse_Ref8U(blob, 8)
    blob.l += 16
    /* CLSID */
    const hlink = parse_Hyperlink(blob, length - 24)
    return [ref, hlink]
}

/* 2.4.141 */
export const parse_HLinkTooltip = function (blob, length) {
    const end = blob.l + length
    blob.read_shift(2)
    const ref = parse_Ref8U(blob, 8)
    let wzTooltip = blob.read_shift((length - 10) / 2, 'dbcs-cont')
    wzTooltip = wzTooltip.replace(chr0, '')
    return [ref, wzTooltip]
}

/* 2.4.63 */
export function parse_Country(blob, length) {
    const o = []
    let d
    d = blob.read_shift(2)
    o[0] = CountryEnum[d] || d
    d = blob.read_shift(2)
    o[1] = CountryEnum[d] || d
    return o
}

/* 2.4.50 ClrtClient */
export function parse_ClrtClient(blob, length) {
    let ccv = blob.read_shift(2)
    const o = []
    while (ccv-- > 0) o.push(parse_LongRGB(blob, 8))
    return o
}

/* 2.4.188 */
export function parse_Palette(blob, length) {
    let ccv = blob.read_shift(2)
    const o = []
    while (ccv-- > 0) o.push(parse_LongRGB(blob, 8))
    return o
}

/* 2.4.354 */
export function parse_XFCRC(blob, length) {
    blob.l += 2
    const o = {cxfs: 0, crc: 0}
    o.cxfs = blob.read_shift(2)
    o.crc = blob.read_shift(4)
    return o
}

/* 2.4.53 TODO: parse flags */
/* [MS-XLSB] 2.4.323 TODO: parse flags */
export function parse_ColInfo(blob, length, opts) {
    if (!opts.cellStyles) return parsenoop(blob, length)
    const w = opts && opts.biff >= 12 ? 4 : 2
    const colFirst = blob.read_shift(w)
    const colLast = blob.read_shift(w)
    const coldx = blob.read_shift(w)
    const ixfe = blob.read_shift(w)
    const flags = blob.read_shift(2)
    if (w == 2) blob.l += 2
    return {s: colFirst, e: colLast, w: coldx, ixfe, flags}
}

/* 2.4.257 */
export function parse_Setup(blob, length, opts) {
    const o = {}
    blob.l += 16
    o.header = parse_Xnum(blob, 8)
    o.footer = parse_Xnum(blob, 8)
    blob.l += 2
    return o
}

/* 2.4.261 */
export function parse_ShtProps(blob, length, opts) {
    const def = {area: false}
    if (opts.biff != 5) {
        blob.l += length
        return def
    }
    const d = blob.read_shift(1)
    blob.l += 3
    if (d & 0x10) def.area = true
    return def
}

export const parse_Style = parsenoop
export const parse_StyleExt = parsenoop

export const parse_Window2 = parsenoop

export const parse_Backup = parsebool
/* 2.4.14 */
export const parse_Blank = parse_XLSCell
/* 2.4.20 Just the cell */
export const parse_BottomMargin = parse_Xnum
/* 2.4.27 */
export const parse_BuiltInFnGroupCount = parseuint16
/* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */
export const parse_CalcCount = parseuint16
/* 2.4.31 #Iterations */
export const parse_CalcDelta = parse_Xnum
/* 2.4.32 */
export const parse_CalcIter = parsebool
/* 2.4.33 1=iterative calc */
export const parse_CalcMode = parseuint16
/* 2.4.34 0=manual, 1=auto (def), 2=table */
export const parse_CalcPrecision = parsebool
/* 2.4.35 */
export const parse_CalcRefMode = parsenoop2
/* 2.4.36 */
export const parse_CalcSaveRecalc = parsebool
/* 2.4.37 */
export const parse_CodePage = parseuint16
/* 2.4.52 */
export const parse_Compat12 = parsebool
/* 2.4.54 true = no compatibility check */
export const parse_Date1904 = parsebool
/* 2.4.77 - 1=1904,0=1900 */
export const parse_DefColWidth = parseuint16
/* 2.4.89 */
export const parse_DSF = parsenoop2
/* 2.4.94 -- MUST be ignored */
export const parse_EntExU2 = parsenoop2
/* 2.4.102 -- Explicitly says to ignore */
export const parse_EOF = parsenoop2
/* 2.4.103 */
export const parse_Excel9File = parsenoop2
/* 2.4.104 -- Optional and unused */
export const parse_FeatHdr = parsenoop2
/* 2.4.112 */
export const parse_FontX = parseuint16
/* 2.4.123 */
export const parse_Footer = parse_XLHeaderFooter
/* 2.4.124 */
export const parse_GridSet = parseuint16
/* 2.4.132, =1 */
export const parse_HCenter = parsebool
/* 2.4.135 sheet centered horizontal on print */
export const parse_Header = parse_XLHeaderFooter
/* 2.4.136 */
export const parse_HideObj = parse_HideObjEnum
/* 2.4.139 */
export const parse_InterfaceEnd = parsenoop2
/* 2.4.145 -- noop */
export const parse_LeftMargin = parse_Xnum
/* 2.4.151 */
export const parse_Mms = parsenoop2
/* 2.4.169 -- Explicitly says to ignore */
export const parse_ObjProtect = parsebool
/* 2.4.183 -- must be 1 if present */
export const parse_Password = parseuint16
/* 2.4.191 */
export const parse_PrintGrid = parsebool
/* 2.4.202 */
export const parse_PrintRowCol = parsebool
/* 2.4.203 */
export const parse_PrintSize = parseuint16
/* 2.4.204 0:3 */
export const parse_Prot4Rev = parsebool
/* 2.4.205 */
export const parse_Prot4RevPass = parseuint16
/* 2.4.206 */
export const parse_Protect = parsebool
/* 2.4.207 */
export const parse_RefreshAll = parsebool
/* 2.4.217 -- must be 0 if not template */
export const parse_RightMargin = parse_Xnum
/* 2.4.219 */
export const parse_RRTabId = parseuint16a
/* 2.4.241 */
export const parse_ScenarioProtect = parsebool
/* 2.4.245 */
export const parse_Scl = parseuint16a
/* 2.4.247 num, den */
export const parse_String = parse_XLUnicodeString
/* 2.4.268 */
export const parse_SxBool = parsebool
/* 2.4.274 */
export const parse_TopMargin = parse_Xnum
/* 2.4.328 */
export const parse_UsesELFs = parsebool
/* 2.4.337 -- should be 0 */
export const parse_VCenter = parsebool
/* 2.4.342 */
export const parse_WinProtect = parsebool
/* 2.4.347 */
export const parse_WriteProtect = parsenoop
/* 2.4.350 empty record */

/* ---- */
export const parse_VerticalPageBreaks = parsenoop
export const parse_HorizontalPageBreaks = parsenoop
export const parse_Selection = parsenoop
export const parse_Continue = parsenoop
export const parse_Pane = parsenoop
export const parse_Pls = parsenoop
export const parse_DCon = parsenoop
export const parse_DConRef = parsenoop
export const parse_DConName = parsenoop
export const parse_XCT = parsenoop
export const parse_CRN = parsenoop
export const parse_FileSharing = parsenoop
export const parse_Uncalced = parsenoop
export const parse_Template = parsenoop
export const parse_Intl = parsenoop
export const parse_WsBool = parsenoop
export const parse_Sort = parsenoop
export const parse_Sync = parsenoop
export const parse_LPr = parsenoop
export const parse_DxGCol = parsenoop
export const parse_FnGroupName = parsenoop
export const parse_FilterMode = parsenoop
export const parse_AutoFilterInfo = parsenoop
export const parse_AutoFilter = parsenoop
export const parse_ScenMan = parsenoop
export const parse_SCENARIO = parsenoop
export const parse_SxView = parsenoop
export const parse_Sxvd = parsenoop
export const parse_SXVI = parsenoop
export const parse_SxIvd = parsenoop
export const parse_SXLI = parsenoop
export const parse_SXPI = parsenoop
export const parse_DocRoute = parsenoop
export const parse_RecipName = parsenoop
export const parse_SXDI = parsenoop
export const parse_SXDB = parsenoop
export const parse_SXFDB = parsenoop
export const parse_SXDBB = parsenoop
export const parse_SXNum = parsenoop
export const parse_SxErr = parsenoop
export const parse_SXInt = parsenoop
export const parse_SXString = parsenoop
export const parse_SXDtr = parsenoop
export const parse_SxNil = parsenoop
export const parse_SXTbl = parsenoop
export const parse_SXTBRGIITM = parsenoop
export const parse_SxTbpg = parsenoop
export const parse_ObProj = parsenoop
export const parse_SXStreamID = parsenoop
export const parse_DBCell = parsenoop
export const parse_SXRng = parsenoop
export const parse_SxIsxoper = parsenoop
export const parse_BookBool = parsenoop
export const parse_DbOrParamQry = parsenoop
export const parse_OleObjectSize = parsenoop
export const parse_SXVS = parsenoop
export const parse_BkHim = parsenoop
export const parse_MsoDrawingGroup = parsenoop
export const parse_MsoDrawing = parsenoop
export const parse_MsoDrawingSelection = parsenoop
export const parse_PhoneticInfo = parsenoop
export const parse_SxRule = parsenoop
export const parse_SXEx = parsenoop
export const parse_SxFilt = parsenoop
export const parse_SxDXF = parsenoop
export const parse_SxItm = parsenoop
export const parse_SxName = parsenoop
export const parse_SxSelect = parsenoop
export const parse_SXPair = parsenoop
export const parse_SxFmla = parsenoop
export const parse_SxFormat = parsenoop
export const parse_SXVDEx = parsenoop
export const parse_SXFormula = parsenoop
export const parse_SXDBEx = parsenoop
export const parse_RRDInsDel = parsenoop
export const parse_RRDHead = parsenoop
export const parse_RRDChgCell = parsenoop
export const parse_RRDRenSheet = parsenoop
export const parse_RRSort = parsenoop
export const parse_RRDMove = parsenoop
export const parse_RRFormat = parsenoop
export const parse_RRAutoFmt = parsenoop
export const parse_RRInsertSh = parsenoop
export const parse_RRDMoveBegin = parsenoop
export const parse_RRDMoveEnd = parsenoop
export const parse_RRDInsDelBegin = parsenoop
export const parse_RRDInsDelEnd = parsenoop
export const parse_RRDConflict = parsenoop
export const parse_RRDDefName = parsenoop
export const parse_RRDRstEtxp = parsenoop
export const parse_LRng = parsenoop
export const parse_CUsr = parsenoop
export const parse_CbUsr = parsenoop
export const parse_UsrInfo = parsenoop
export const parse_UsrExcl = parsenoop
export const parse_FileLock = parsenoop
export const parse_RRDInfo = parsenoop
export const parse_BCUsrs = parsenoop
export const parse_UsrChk = parsenoop
export const parse_UserBView = parsenoop
export const parse_UserSViewBegin = parsenoop // overloaded
export const parse_UserSViewEnd = parsenoop
export const parse_RRDUserView = parsenoop
export const parse_Qsi = parsenoop
export const parse_CondFmt = parsenoop
export const parse_CF = parsenoop
export const parse_DVal = parsenoop
export const parse_DConBin = parsenoop
export const parse_Lel = parsenoop
export const parse_XLSCodeName = parse_XLUnicodeString
export const parse_SXFDBType = parsenoop
export const parse_ObNoMacros = parsenoop
export const parse_Dv = parsenoop
export const parse_Index = parsenoop
export const parse_Table = parsenoop
export const parse_BigName = parsenoop
export const parse_ContinueBigName = parsenoop
export const parse_WebPub = parsenoop
export const parse_QsiSXTag = parsenoop
export const parse_DBQueryExt = parsenoop
export const parse_ExtString = parsenoop
export const parse_TxtQry = parsenoop
export const parse_Qsir = parsenoop
export const parse_Qsif = parsenoop
export const parse_RRDTQSIF = parsenoop
export const parse_OleDbConn = parsenoop
export const parse_WOpt = parsenoop
export const parse_SXViewEx = parsenoop
export const parse_SXTH = parsenoop
export const parse_SXPIEx = parsenoop
export const parse_SXVDTEx = parsenoop
export const parse_SXViewEx9 = parsenoop
export const parse_ContinueFrt = parsenoop
export const parse_RealTimeData = parsenoop
export const parse_ChartFrtInfo = parsenoop
export const parse_FrtWrapper = parsenoop
export const parse_StartBlock = parsenoop
export const parse_EndBlock = parsenoop
export const parse_StartObject = parsenoop
export const parse_EndObject = parsenoop
export const parse_CatLab = parsenoop
export const parse_YMult = parsenoop
export const parse_SXViewLink = parsenoop
export const parse_PivotChartBits = parsenoop
export const parse_FrtFontList = parsenoop
export const parse_SheetExt = parsenoop
export const parse_BookExt = parsenoop
export const parse_SXAddl = parsenoop
export const parse_CrErr = parsenoop
export const parse_HFPicture = parsenoop
export const parse_Feat = parsenoop
export const parse_DataLabExt = parsenoop
export const parse_DataLabExtContents = parsenoop
export const parse_CellWatch = parsenoop
export const parse_FeatHdr11 = parsenoop
export const parse_Feature11 = parsenoop
export const parse_DropDownObjIds = parsenoop
export const parse_ContinueFrt11 = parsenoop
export const parse_DConn = parsenoop
export const parse_List12 = parsenoop
export const parse_Feature12 = parsenoop
export const parse_CondFmt12 = parsenoop
export const parse_CF12 = parsenoop
export const parse_CFEx = parsenoop
export const parse_AutoFilter12 = parsenoop
export const parse_ContinueFrt12 = parsenoop
export const parse_MDTInfo = parsenoop
export const parse_MDXStr = parsenoop
export const parse_MDXTuple = parsenoop
export const parse_MDXSet = parsenoop
export const parse_MDXProp = parsenoop
export const parse_MDXKPI = parsenoop
export const parse_MDB = parsenoop
export const parse_PLV = parsenoop
export const parse_DXF = parsenoop
export const parse_TableStyles = parsenoop
export const parse_TableStyle = parsenoop
export const parse_TableStyleElement = parsenoop
export const parse_NamePublish = parsenoop
export const parse_SortData = parsenoop
export const parse_GUIDTypeLib = parsenoop
export const parse_FnGrp12 = parsenoop
export const parse_NameFnGrp12 = parsenoop
export const parse_HeaderFooter = parsenoop
export const parse_CrtLayout12 = parsenoop
export const parse_CrtMlFrt = parsenoop
export const parse_CrtMlFrtContinue = parsenoop
export const parse_ShapePropsStream = parsenoop
export const parse_TextPropsStream = parsenoop
export const parse_RichTextStream = parsenoop
export const parse_CrtLayout12A = parsenoop
export const parse_Units = parsenoop
export const parse_Chart = parsenoop
export const parse_Series = parsenoop
export const parse_DataFormat = parsenoop
export const parse_LineFormat = parsenoop
export const parse_MarkerFormat = parsenoop
export const parse_AreaFormat = parsenoop
export const parse_PieFormat = parsenoop
export const parse_AttachedLabel = parsenoop
export const parse_SeriesText = parsenoop
export const parse_ChartFormat = parsenoop
export const parse_Legend = parsenoop
export const parse_SeriesList = parsenoop
export const parse_Bar = parsenoop
export const parse_Line = parsenoop
export const parse_Pie = parsenoop
export const parse_Area = parsenoop
export const parse_Scatter = parsenoop
export const parse_CrtLine = parsenoop
export const parse_Axis = parsenoop
export const parse_Tick = parsenoop
export const parse_ValueRange = parsenoop
export const parse_CatSerRange = parsenoop
export const parse_AxisLine = parsenoop
export const parse_CrtLink = parsenoop
export const parse_DefaultText = parsenoop
export const parse_Text = parsenoop
export const parse_ObjectLink = parsenoop
export const parse_Frame = parsenoop
export const parse_Begin = parsenoop
export const parse_End = parsenoop
export const parse_PlotArea = parsenoop
export const parse_Chart3d = parsenoop
export const parse_PicF = parsenoop
export const parse_DropBar = parsenoop
export const parse_Radar = parsenoop
export const parse_Surf = parsenoop
export const parse_RadarArea = parsenoop
export const parse_AxisParent = parsenoop
export const parse_LegendException = parsenoop
export const parse_SerToCrt = parsenoop
export const parse_AxesUsed = parsenoop
export const parse_SBaseRef = parsenoop
export const parse_SerParent = parsenoop
export const parse_SerAuxTrend = parsenoop
export const parse_IFmtRecord = parsenoop
export const parse_Pos = parsenoop
export const parse_AlRuns = parsenoop
export const parse_BRAI = parsenoop
export const parse_SerAuxErrBar = parsenoop
export const parse_SerFmt = parsenoop
export const parse_Chart3DBarShape = parsenoop
export const parse_Fbi = parsenoop
export const parse_BopPop = parsenoop
export const parse_AxcExt = parsenoop
export const parse_Dat = parsenoop
export const parse_PlotGrowth = parsenoop
export const parse_SIIndex = parsenoop
export const parse_GelFrame = parsenoop
export const parse_BopPopCustom = parsenoop
export const parse_Fbi2 = parsenoop

/* --- Specific to versions before BIFF8 --- */
export function parse_BIFF5String(blob) {
    const len = blob.read_shift(1)
    return blob.read_shift(len, 'sbcs-cont')
}

/* BIFF2_??? where ??? is the name from [XLS] */
export function parse_BIFF2STR(blob, length, opts) {
    const cell = parse_XLSCell(blob, 6)
    ++blob.l
    const str = parse_XLUnicodeString2(blob, length - 7, opts)
    cell.t = 'str'
    cell.val = str
    return cell
}

export function parse_BIFF2NUM(blob, length, opts) {
    const cell = parse_XLSCell(blob, 6)
    ++blob.l
    const num = parse_Xnum(blob, 8)
    cell.t = 'n'
    cell.val = num
    return cell
}

export function parse_BIFF2INT(blob, length) {
    const cell = parse_XLSCell(blob, 6)
    ++blob.l
    const num = blob.read_shift(2)
    cell.t = 'n'
    cell.val = num
    return cell
}

export function parse_BIFF2STRING(blob, length) {
    const cch = blob.read_shift(1)
    if (cch === 0) {
        blob.l++
        return ''
    }
    return blob.read_shift(cch, 'sbcs-cont')
}

/* TODO: convert to BIFF8 font struct */
export function parse_BIFF2FONTXTRA(blob, length) {
    blob.l += 6 // unknown
    blob.l += 2 // font weight "bls"
    blob.l += 1 // charset
    blob.l += 3 // unknown
    blob.l += 1 // font family
    blob.l += length - 9
}

/* TODO: parse rich text runs */
export function parse_RString(blob, length, opts) {
    const end = blob.l + length
    const cell = parse_XLSCell(blob, 6)
    const cch = blob.read_shift(2)
    const str = parse_XLUnicodeStringNoCch(blob, cch, opts)
    blob.l = end
    cell.t = 'str'
    cell.val = str
    return cell
}
