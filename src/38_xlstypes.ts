import { current_codepage, setCurrentCodepage } from './02_codepage'
import { chr0, chr1 } from './05_buf'
import * as CFB from './18_cfb'
import { parseDate } from './20_jsutils'
import { parsenoop, prep_blob } from './23_binutils'
import { VT_CUSTOM, VT_I2, VT_I4, VT_USTR, VT_VARIANT } from './29_xlsenum'

/* [MS-DTYP] 2.3.3 FILETIME */
/* [MS-OLEDS] 2.1.3 FILETIME (Packet Version) */
/* [MS-OLEPS] 2.8 FILETIME (Packet Version) */
function parse_FILETIME(blob) {
    const dwLowDateTime = blob.read_shift(4)
    const dwHighDateTime = blob.read_shift(4)
    return new Date((dwHighDateTime / 1e7 * Math.pow(2, 32) + dwLowDateTime / 1e7 - 11644473600) * 1000).toISOString()
        .replace(/\.000/, '')
}

/* [MS-OSHARED] 2.3.3.1.4 Lpstr */
function parse_lpstr(blob, type?, pad?) {
    const str = blob.read_shift(0, 'lpstr')
    if (pad) {
        blob.l += 4 - (str.length + 1 & 3) & 3
    }
    return str
}

/* [MS-OSHARED] 2.3.3.1.6 Lpwstr */
function parse_lpwstr(blob, type?, pad?) {
    const str = blob.read_shift(0, 'lpwstr')
    if (pad) {
        blob.l += 4 - (str.length + 1 & 3) & 3
    }
    return str
}

/* [MS-OSHARED] 2.3.3.1.11 VtString */
/* [MS-OSHARED] 2.3.3.1.12 VtUnalignedString */
function parse_VtStringBase(blob, stringType, pad) {
    if (stringType === 0x1F /*VT_LPWSTR*/) {
        return parse_lpwstr(blob)
    }
    return parse_lpstr(blob, stringType, pad)
}

function parse_VtString(blob, t, pad?) {
    return parse_VtStringBase(blob, t, pad === false ? 0 : 4)
}
function parse_VtUnalignedString(blob, t) {
    if (!t) {
        throw new Error('dafuq?')
    }
    return parse_VtStringBase(blob, t, 0)
}

/* [MS-OSHARED] 2.3.3.1.9 VtVecUnalignedLpstrValue */
function parse_VtVecUnalignedLpstrValue(blob) {
    const length = blob.read_shift(4)
    const ret = []
    for (let i = 0; i != length; ++i) {
        ret[i] = blob.read_shift(0, 'lpstr')
    }
    return ret
}

/* [MS-OSHARED] 2.3.3.1.10 VtVecUnalignedLpstr */
function parse_VtVecUnalignedLpstr(blob) {
    return parse_VtVecUnalignedLpstrValue(blob)
}

/* [MS-OSHARED] 2.3.3.1.13 VtHeadingPair */
function parse_VtHeadingPair(blob) {
    const headingString = parse_TypedPropertyValue(blob, VT_USTR)
    const headerParts = parse_TypedPropertyValue(blob, VT_I4)
    return [headingString, headerParts]
}

/* [MS-OSHARED] 2.3.3.1.14 VtVecHeadingPairValue */
function parse_VtVecHeadingPairValue(blob) {
    const cElements = blob.read_shift(4)
    const out = []
    for (let i = 0; i != cElements / 2; ++i) {
        out.push(parse_VtHeadingPair(blob))
    }
    return out
}

/* [MS-OSHARED] 2.3.3.1.15 VtVecHeadingPair */
function parse_VtVecHeadingPair(blob) {
    // NOTE: When invoked, wType & padding were already consumed
    return parse_VtVecHeadingPairValue(blob)
}

/* [MS-OLEPS] 2.18.1 Dictionary (uses 2.17, 2.16) */
function parse_dictionary(blob, CodePage) {
    const cnt = blob.read_shift(4)
    const dict /*:{[number]:string}*/ = {}
    /*:any*/
    for (let j = 0; j != cnt; ++j) {
        const pid = blob.read_shift(4)
        const len = blob.read_shift(4)
        dict[pid] = blob.read_shift(len, CodePage === 0x4B0 ? 'utf16le' : 'utf8').replace(chr0, '').replace(chr1, '!')
    }
    if (blob.l & 3) {
        blob.l = blob.l >> 2 + 1 << 2
    }
    return dict
}

/* [MS-OLEPS] 2.9 BLOB */
function parse_BLOB(blob) {
    const size = blob.read_shift(4)
    const bytes = blob.slice(blob.l, blob.l + size)
    if ((size & 3) > 0) {
        blob.l += 4 - (size & 3) & 3
    }
    return bytes
}

/* [MS-OLEPS] 2.11 ClipboardData */
function parse_ClipboardData(blob) {
    // TODO
    const o = {}
    o.Size = blob.read_shift(4)
    //o.Format = blob.read_shift(4);
    blob.l += o.Size
    return o
}

/* [MS-OLEPS] 2.14 Vector and Array Property Types */
function parse_VtVector(blob, cb) {
}
/* [MS-OLEPS] 2.14.2 VectorHeader */
/*	var Length = blob.read_shift(4);
 var o = [];
 for(var i = 0; i != Length; ++i) {
 o.push(cb(blob));
 }
 return o;*/


/* [MS-OLEPS] 2.15 TypedPropertyValue */
function parse_TypedPropertyValue(blob, type, _opts?) {
    const t = blob.read_shift(2)
    let ret
    const opts = _opts || {}
    blob.l += 2
    if (type !== VT_VARIANT) {
        if (t !== type && !VT_CUSTOM.includes(type)) {
            throw new Error(`Expected type ${type} saw ${t}`)
        }
    }
    switch (type === VT_VARIANT ? t : type) {
        case 0x02 /*VT_I2*/
        :
            ret = blob.read_shift(2, 'i')
            if (!opts.raw) {
                blob.l += 2
            }
            return ret
        case 0x03 /*VT_I4*/
        :
            ret = blob.read_shift(4, 'i')
            return ret
        case 0x0B /*VT_BOOL*/
        :
            return blob.read_shift(4) !== 0x0
        case 0x13 /*VT_UI4*/
        :
            ret = blob.read_shift(4)
            return ret
        case 0x1E /*VT_LPSTR*/
        :
            return parse_lpstr(blob, t, 4).replace(chr0, '')
        case 0x1F /*VT_LPWSTR*/
        :
            return parse_lpwstr(blob)
        case 0x40 /*VT_FILETIME*/
        :
            return parse_FILETIME(blob)
        case 0x41 /*VT_BLOB*/
        :
            return parse_BLOB(blob)
        case 0x47 /*VT_CF*/
        :
            return parse_ClipboardData(blob)
        case 0x50 /*VT_STRING*/
        :
            return parse_VtString(blob, t, !opts.raw && 4).replace(chr0, '')
        case 0x51 /*VT_USTR*/
        :
            return parse_VtUnalignedString(blob, t, 4).replace(chr0, '')
        case 0x100C /*VT_VECTOR|VT_VARIANT*/
        :
            return parse_VtVecHeadingPair(blob)
        case 0x101E /*VT_LPSTR*/
        :
            return parse_VtVecUnalignedLpstr(blob)
        default:
            throw new Error(`TypedPropertyValue unrecognized type ${type} ${t}`)
    }
}
/* [MS-OLEPS] 2.14.2 VectorHeader */
/*function parse_VTVectorVariant(blob) {
 var Length = blob.read_shift(4);

 if(Length & 1 !== 0) throw new Error("VectorHeader Length=" + Length + " must be even");
 var o = [];
 for(var i = 0; i != Length; ++i) {
 o.push(parse_TypedPropertyValue(blob, VT_VARIANT));
 }
 return o;
 }*/

/* [MS-OLEPS] 2.20 PropertySet */
function parse_PropertySet(blob, PIDSI) {
    const start_addr = blob.l
    const size = blob.read_shift(4)
    const NumProps = blob.read_shift(4)
    const Props = []
    let i = 0
    let CodePage = 0
    let Dictionary = -1
    let DictObj /*:{[number]:string}*/ = {}
    /*:any*/
    for (i = 0; i != NumProps; ++i) {
        const PropID = blob.read_shift(4)
        const Offset = blob.read_shift(4)
        Props[i] = [PropID, Offset + start_addr]
    }
    const PropH = {}
    for (i = 0; i != NumProps; ++i) {
        if (blob.l !== Props[i][1]) {
            let fail = true
            if (i > 0 && PIDSI) {
                switch (PIDSI[Props[i - 1][0]].t) {
                    case 0x02: /*VT_I2*/

                        if (blob.l + 2 === Props[i][1]) {
                            blob.l += 2
                            fail = false
                        }
                        break
                    case 0x50: /*VT_STRING*/

                        if (blob.l <= Props[i][1]) {
                            blob.l = Props[i][1]
                            fail = false
                        }
                        break
                    case 0x100C: /*VT_VECTOR|VT_VARIANT*/

                        if (blob.l <= Props[i][1]) {
                            blob.l = Props[i][1]
                            fail = false
                        }
                        break
                }
            }
            if (!PIDSI && blob.l <= Props[i][1]) {
                fail = false
                blob.l = Props[i][1]
            }
            if (fail) {
                throw new Error(`Read Error: Expected address ${Props[i][1]} at ${blob.l} :${i}`)
            }
        }
        if (PIDSI) {
            const piddsi = PIDSI[Props[i][0]]
            PropH[piddsi.n] = parse_TypedPropertyValue(blob, piddsi.t, { raw: true })
            if (piddsi.p === 'version') {
                PropH[piddsi.n] = `${String(PropH[piddsi.n] >> 16)}.${String(PropH[piddsi.n] & 0xFFFF)}`
            }
            if (piddsi.n == 'CodePage') {
                switch (PropH[piddsi.n]) {
                    case 0:
                        PropH[piddsi.n] = 1252
                    /* falls through */
                    case 10000: // OSX Roman
                    case 1252: // Windows Latin

                    case 874: // SB Windows Thai
                    case 1250: // SB Windows Central Europe
                    case 1251: // SB Windows Cyrillic
                    case 1253: // SB Windows Greek
                    case 1254: // SB Windows Turkish
                    case 1255: // SB Windows Hebrew
                    case 1256: // SB Windows Arabic
                    case 1257: // SB Windows Baltic
                    case 1258: // SB Windows Vietnam

                    case 932: // DB Windows Japanese Shift-JIS
                    case 936: // DB Windows Simplified Chinese GBK
                    case 949: // DB Windows Korean
                    case 950: // DB Windows Traditional Chinese Big5

                    case 1200: // UTF16LE
                    case 1201: // UTF16BE
                    case 65000:
                    case -536: // UTF-7
                    case 65001:
                    case -535:
                        // UTF-8
                        setCurrentCodepage(CodePage = PropH[piddsi.n])
                        break
                    default:
                        throw new Error(`Unsupported CodePage: ${PropH[piddsi.n]}`)
                }
            }
        } else {
            if (Props[i][0] === 0x1) {
                CodePage = PropH.CodePage = parse_TypedPropertyValue(blob, VT_I2)
                setCurrentCodepage(CodePage)
                if (Dictionary !== -1) {
                    const oldpos = blob.l
                    blob.l = Props[Dictionary][1]
                    DictObj = parse_dictionary(blob, CodePage)
                    blob.l = oldpos
                }
            } else if (Props[i][0] === 0) {
                if (CodePage === 0) {
                    Dictionary = i
                    blob.l = Props[i + 1][1]
                    continue
                }
                DictObj = parse_dictionary(blob, CodePage)
            } else {
                const name = DictObj[Props[i][0]]
                let val
                /* [MS-OSHARED] 2.3.3.2.3.1.2 + PROPVARIANT */
                switch (blob[blob.l]) {
                    case 0x41: /*VT_BLOB*/

                        blob.l += 4
                        val = parse_BLOB(blob)
                        break
                    case 0x1E: /*VT_LPSTR*/

                        blob.l += 4
                        val = parse_VtString(blob, blob[blob.l - 4])
                        break
                    case 0x1F: /*VT_LPWSTR*/

                        blob.l += 4
                        val = parse_VtString(blob, blob[blob.l - 4])
                        break
                    case 0x03: /*VT_I4*/

                        blob.l += 4
                        val = blob.read_shift(4, 'i')
                        break
                    case 0x13: /*VT_UI4*/

                        blob.l += 4
                        val = blob.read_shift(4)
                        break
                    case 0x05: /*VT_R8*/

                        blob.l += 4
                        val = blob.read_shift(8, 'f')
                        break
                    case 0x0B: /*VT_BOOL*/

                        blob.l += 4
                        val = parsebool(blob, 4)
                        break
                    case 0x40: /*VT_FILETIME*/

                        blob.l += 4
                        val = parseDate(parse_FILETIME(blob))
                        break
                    default:
                        throw new Error(`unparsed value: ${blob[blob.l]}`)
                }
                PropH[name] = val
            }
        }
    }
    blob.l = start_addr + size
    /* step ahead to skip padding */
    return PropH
}

/* [MS-OLEPS] 2.21 PropertySetStream */
export function parse_PropertySetStream(file, PIDSI) {
    const blob = file.content
    prep_blob(blob, 0)

    let NumSets
    let FMTID0
    let FMTID1
    let Offset0
    let Offset1 = 0
    blob.chk('feff', 'Byte Order: ')

    const vers = blob.read_shift(2) // TODO: check version
    const SystemIdentifier = blob.read_shift(4)
    blob.chk(CFB.utils.consts.HEADER_CLSID, 'CLSID: ')
    NumSets = blob.read_shift(4)
    if (NumSets !== 1 && NumSets !== 2) {
        throw new Error(`Unrecognized #Sets: ${NumSets}`)
    }
    FMTID0 = blob.read_shift(16)
    Offset0 = blob.read_shift(4)

    if (NumSets === 1 && Offset0 !== blob.l) {
        throw new Error(`Length mismatch: ${Offset0} !== ${blob.l}`)
    } else if (NumSets === 2) {
        FMTID1 = blob.read_shift(16)
        Offset1 = blob.read_shift(4)
    }
    const PSet0 = parse_PropertySet(blob, PIDSI)

    const rval = { SystemIdentifier }
    /*:any*/
    for (let y in PSet0) {
        rval[y] = PSet0[y]
    }
    //rval.blob = blob;
    rval.FMTID = FMTID0
    //rval.PSet0 = PSet0;
    if (NumSets === 1) {
        return rval
    }
    if (blob.l !== Offset1) {
        throw new Error(`Length mismatch 2: ${blob.l} !== ${Offset1}`)
    }
    let PSet1
    try {
        PSet1 = parse_PropertySet(blob, null)
    } catch (e) {
    }
    for (let y in PSet1) {
        rval[y] = PSet1[y]
    }
    rval.FMTID = [FMTID0, FMTID1] // TODO: verify FMTID0/1
    return rval
}

export function parsenoop2(blob, length) {
    blob.read_shift(length)
    return null
}

function parslurp(blob, length, cb) {
    const arr = []
    const target = blob.l + length
    while (blob.l < target) {
        arr.push(cb(blob, target - blob.l))
    }
    if (target !== blob.l) {
        throw new Error('Slurp error')
    }
    return arr
}

export function parsebool(blob, length) {
    return blob.read_shift(length) === 0x1
}

export function parseuint16(blob) {
    return blob.read_shift(2, 'u')
}

export function parseuint16a(blob, length) {
    return parslurp(blob, length, parseuint16)
}

/* --- 2.5 Structures --- */

/* [MS-XLS] 2.5.14 Boolean */
const parse_Boolean = parsebool

/* [MS-XLS] 2.5.10 Bes (boolean or error) */
export function parse_Bes(blob) {
    const v = blob.read_shift(1)
    const t = blob.read_shift(1)
    return t === 0x01 ? v : v === 0x01
}

/* [MS-XLS] 2.5.240 ShortXLUnicodeString */
export function parse_ShortXLUnicodeString(blob, length, opts) {
    const cch = blob.read_shift(opts && opts.biff >= 12 ? 2 : 1)
    let width = 1
    let encoding = 'sbcs-cont'
    const cp = current_codepage
    if (opts && opts.biff >= 8) {
        setCurrentCodepage(1200)
    }
    if (!opts || opts.biff == 8) {
        const fHighByte = blob.read_shift(1)
        if (fHighByte) {
            width = 2
            encoding = 'dbcs-cont'
        }
    } else if (opts.biff == 12) {
        width = 2
        encoding = 'wstr'
    }
    const o = cch ? blob.read_shift(cch, encoding) : ''
    setCurrentCodepage(cp)
    return o
}

/* 2.5.293 XLUnicodeRichExtendedString */
export function parse_XLUnicodeRichExtendedString(blob) {
    const cp = current_codepage
    setCurrentCodepage(1200)
    const cch = blob.read_shift(2)
    const flags = blob.read_shift(1)
    const fHighByte = flags & 0x1
    const fExtSt = flags & 0x4
    const fRichSt = flags & 0x8
    const width = 1 + (flags & 0x1) // 0x0 -> utf8, 0x1 -> dbcs
    let cRun = 0
    let cbExtRst
    const z = {}
    if (fRichSt) {
        cRun = blob.read_shift(2)
    }
    if (fExtSt) {
        cbExtRst = blob.read_shift(4)
    }
    const encoding = flags & 0x1 ? 'dbcs-cont' : 'sbcs-cont'
    const msg = cch === 0 ? '' : blob.read_shift(cch, encoding)
    if (fRichSt) {
        blob.l += 4 * cRun
    } //TODO: parse this
    if (fExtSt) {
        blob.l += cbExtRst
    } //TODO: parse this
    z.t = msg
    if (!fRichSt) {
        z.raw = `<t>${z.t}</t>`
        z.r = z.t
    }
    setCurrentCodepage(cp)
    return z
}

/* 2.5.296 XLUnicodeStringNoCch */
export function parse_XLUnicodeStringNoCch(blob, cch, opts?) {
    let retval
    if (opts) {
        if (opts.biff >= 2 && opts.biff <= 5) {
            return blob.read_shift(cch, 'sbcs-cont')
        }
        if (opts.biff >= 12) {
            return blob.read_shift(cch, 'dbcs-cont')
        }
    }
    const fHighByte = blob.read_shift(1)
    if (fHighByte === 0) {
        retval = blob.read_shift(cch, 'sbcs-cont')
    } else {
        retval = blob.read_shift(cch, 'dbcs-cont')
    }
    return retval
}

/* 2.5.294 XLUnicodeString */
export function parse_XLUnicodeString(blob, length, opts) {
    const cch = blob.read_shift(opts && opts.biff == 2 ? 1 : 2)
    if (cch === 0) {
        blob.l++
        return ''
    }
    return parse_XLUnicodeStringNoCch(blob, cch, opts)
}
/* BIFF5 override */
export function parse_XLUnicodeString2(blob, length, opts) {
    if (opts.biff > 5) {
        return parse_XLUnicodeString(blob, length, opts)
    }
    const cch = blob.read_shift(1)
    if (cch === 0) {
        blob.l++
        return ''
    }
    return blob.read_shift(cch, 'sbcs-cont')
}

/* [MS-XLS] 2.5.61 ControlInfo */
export const parse_ControlInfo = parsenoop

/* [MS-OSHARED] 2.3.7.6 URLMoniker TODO: flags */
const parse_URLMoniker = function (blob, length) {
    const len = blob.read_shift(4)
    const start = blob.l
    let extra = false
    if (len > 24) {
        /* look ahead */
        blob.l += len - 24
        if (blob.read_shift(16) === '795881f43b1d7f48af2c825dc4852763') {
            extra = true
        }
        blob.l = start
    }
    const url = blob.read_shift((extra ? len - 24 : len) >> 1, 'utf16le').replace(chr0, '')
    if (extra) {
        blob.l += 24
    }
    return url
}

/* [MS-OSHARED] 2.3.7.8 FileMoniker TODO: all fields */
const parse_FileMoniker = function (blob, length) {
    const cAnti = blob.read_shift(2)
    const ansiLength = blob.read_shift(4)
    const ansiPath = blob.read_shift(ansiLength, 'cstr')
    const endServer = blob.read_shift(2)
    const versionNumber = blob.read_shift(2)
    const cbUnicodePathSize = blob.read_shift(4)
    if (cbUnicodePathSize === 0) {
        return ansiPath.replace(/\\/g, '/')
    }
    const cbUnicodePathBytes = blob.read_shift(4)
    const usKeyValue = blob.read_shift(2)
    const unicodePath = blob.read_shift(cbUnicodePathBytes >> 1, 'utf16le').replace(chr0, '')
    return unicodePath
}

/* [MS-OSHARED] 2.3.7.2 HyperlinkMoniker TODO: all the monikers */
const parse_HyperlinkMoniker = function (blob, length) {
    const clsid = blob.read_shift(16)
    length -= 16
    switch (clsid) {
        case 'e0c9ea79f9bace118c8200aa004ba90b':
            return parse_URLMoniker(blob, length)
        case '0303000000000000c000000000000046':
            return parse_FileMoniker(blob, length)
        default:
            throw new Error(`Unsupported Moniker ${clsid}`)
    }
}

/* [MS-OSHARED] 2.3.7.9 HyperlinkString */
const parse_HyperlinkString = function (blob, length) {
    const len = blob.read_shift(4)
    const o = blob.read_shift(len, 'utf16le').replace(chr0, '')
    return o
}

/* [MS-OSHARED] 2.3.7.1 Hyperlink Object TODO: unify params with XLSX */
export const parse_Hyperlink = function (blob, length) {
    const end = blob.l + length
    const sVer = blob.read_shift(4)
    if (sVer !== 2) {
        throw new Error(`Unrecognized streamVersion: ${sVer}`)
    }
    const flags = blob.read_shift(2)
    blob.l += 2
    let displayName
    let targetFrameName
    let moniker
    let oleMoniker
    let location
    let guid
    let fileTime
    if (flags & 0x0010) {
        displayName = parse_HyperlinkString(blob, end - blob.l)
    }
    if (flags & 0x0080) {
        targetFrameName = parse_HyperlinkString(blob, end - blob.l)
    }
    if ((flags & 0x0101) === 0x0101) {
        moniker = parse_HyperlinkString(blob, end - blob.l)
    }
    if ((flags & 0x0101) === 0x0001) {
        oleMoniker = parse_HyperlinkMoniker(blob, end - blob.l)
    }
    if (flags & 0x0008) {
        location = parse_HyperlinkString(blob, end - blob.l)
    }
    if (flags & 0x0020) {
        guid = blob.read_shift(16)
    }
    if (flags & 0x0040) {
        fileTime = parse_FILETIME(blob, 8)
    }
    blob.l = end
    let target = targetFrameName || moniker || oleMoniker
    if (location) {
        target += `#${location}`
    }
    return { Target: target }
}

/* 2.5.178 LongRGBA */
export function parse_LongRGBA(blob, length) {
    const r = blob.read_shift(1)
    const g = blob.read_shift(1)
    const b = blob.read_shift(1)
    const a = blob.read_shift(1)
    return [r, g, b, a]
}

/* 2.5.177 LongRGB */
export function parse_LongRGB(blob, length) {
    const x = parse_LongRGBA(blob, length)
    x[3] = 0
    return x
}
