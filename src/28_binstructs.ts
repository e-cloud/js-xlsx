/* [MS-XLSB] 2.5.143 */
import { evert_num } from './20_jsutils'
import { __double, __readInt32LE, new_buf } from './23_binutils'

function parse_StrRun(data, length?: number) {
    return { ich: data.read_shift(2), ifnt: data.read_shift(2) }
}

/* [MS-XLSB] 2.1.7.121 */
export function parse_RichStr(data, length: number): XLString {
    const start = data.l
    const flags = data.read_shift(1)
    const str = parse_XLWideString(data)
    const rgsStrRun = []
    const z = { t: str, h: str }

    if ((flags & 1) !== 0) {
        /* fRichStr */
        /* TODO: formatted string */
        const dwSizeStrRun = data.read_shift(4)
        for (let i = 0; i != dwSizeStrRun; ++i) {
            rgsStrRun.push(parse_StrRun(data))
        }
        z.r = rgsStrRun
    } else {
        z.r = [{ ich: 0, ifnt: 0 }]
    }
    //if((flags & 2) !== 0) { /* fExtStr */
    //	/* TODO: phonetic string */
    //}
    data.l = start + length
    return z
}

export function write_RichStr(str: XLString, o?: Block): Block {
    /* TODO: formatted string */
    let _null = false
    if (o == null) {
        _null = true
        o = new_buf(15 + 4 * str.t.length)
    }
    o.write_shift(1, 0)
    write_XLWideString(str.t, o)
    return _null ? o.slice(0, o.l) : o
}

/* [MS-XLSB] 2.5.9 */
export function parse_XLSBCell(data) {
    const col = data.read_shift(4)
    let iStyleRef = data.read_shift(2)
    iStyleRef += data.read_shift(1) << 16
    const fPhShow = data.read_shift(1)
    return { c: col, iStyleRef }
}
export function write_XLSBCell(cell, o ?: Block) {
    if (o == null) {
        o = new_buf(8)
    }
    o.write_shift(-4, cell.c)
    o.write_shift(3, cell.iStyleRef || cell.s)
    o.write_shift(1, 0)
    /* fPhShow */
    return o
}

/* [MS-XLSB] 2.5.21 */
export const parse_XLSBCodeName = parse_XLWideString
export const write_XLSBCodeName = write_XLWideString

/* [MS-XLSB] 2.5.166 */
export function parse_XLNullableWideString(data): string {
    const cchCharacters = data.read_shift(4)
    return cchCharacters === 0 || cchCharacters === 0xFFFFFFFF ? '' : data.read_shift(cchCharacters, 'dbcs')
}
export function write_XLNullableWideString(data: string, o) {
    let _null = false
    if (o == null) {
        _null = true
        o = new_buf(127)
    }
    o.write_shift(4, data.length > 0 ? data.length : 0xFFFFFFFF)
    if (data.length > 0) {
        o.write_shift(0, data, 'dbcs')
    }
    return _null ? o.slice(0, o.l) : o
}

/* [MS-XLSB] 2.5.168 */
export function parse_XLWideString(data): string {
    const cchCharacters = data.read_shift(4)
    return cchCharacters === 0 ? '' : data.read_shift(cchCharacters, 'dbcs')
}

export function write_XLWideString(data: string, o?) {
    let _null = false
    if (o == null) {
        _null = true
        o = new_buf(4 + 2 * data.length)
    }
    o.write_shift(4, data.length)
    if (data.length > 0) {
        o.write_shift(0, data, 'dbcs')
    }
    return _null ? o.slice(0, o.l) : o
}

/* [MS-XLSB] 2.5.165 */
export const parse_XLNameWideString = parse_XLWideString
export const write_XLNameWideString = write_XLWideString

/* [MS-XLSB] 2.5.114 */
export const parse_RelID = parse_XLNullableWideString
export const write_RelID = write_XLNullableWideString

/* [MS-XLSB] 2.5.122 */
/* [MS-XLS] 2.5.217 */
export function parse_RkNumber(data): number {
    const b = data.slice(data.l, data.l + 4)
    const fX100 = b[0] & 1
    const fInt = b[0] & 2
    data.l += 4
    b[0] &= 0xFC // b[0] &= ~3;
    const RK = fInt === 0 ? __double([0, 0, 0, 0, b[0], b[1], b[2], b[3]], 0) : __readInt32LE(b, 0) >> 2
    return fX100 ? RK / 100 : RK
}
export function write_RkNumber(data: number, o) {
    if (o == null) {
        o = new_buf(4)
    }
    let fX100 = 0
    let fInt = 0
    const d100 = data * 100
    if (data == (data | 0) && data >= -(1 << 29) && data < 1 << 29) {
        fInt = 1
    } else if (d100 == (d100 | 0) && d100 >= -(1 << 29) && d100 < 1 << 29) {
        fInt = 1
        fX100 = 1
    }
    if (fInt) {
        o.write_shift(-4, ((fX100 ? d100 : data) << 2) + (fX100 + 2))
    } else {
        throw new Error(`unsupported RkNumber ${data}`)
    } // TODO
}

/* [MS-XLSB] 2.5.117 RfX */
export function parse_RfX(data): Range {
    const cell: Range = { s: {}, e: {} }

    cell.s.r = data.read_shift(4)
    cell.e.r = data.read_shift(4)
    cell.s.c = data.read_shift(4)
    cell.e.c = data.read_shift(4)
    return cell
}

function write_RfX(r: Range, o) {
    if (!o) {
        o = new_buf(16)
    }
    o.write_shift(4, r.s.r)
    o.write_shift(4, r.e.r)
    o.write_shift(4, r.s.c)
    o.write_shift(4, r.e.c)
    return o
}

/* [MS-XLSB] 2.5.153 UncheckedRfX */
export const parse_UncheckedRfX = parse_RfX
export const write_UncheckedRfX = write_RfX

/* [MS-XLSB] 2.5.171 */
/* [MS-XLS] 2.5.342 */
/* TODO: error checking, NaN and Infinity values are not valid Xnum */
export function parse_Xnum(data, length?: number) {
    return data.read_shift(8, 'f')
}
export function write_Xnum(data, o) {
    return (o || new_buf(8)).write_shift(8, data, 'f')
}

/* [MS-XLSB] 2.5.198.2 */
export const BErr = {
    /*::[*/0x00 /*::]*/: '#NULL!',
    /*::[*/0x07 /*::]*/: '#DIV/0!',
    /*::[*/0x0F /*::]*/: '#VALUE!',
    /*::[*/0x17 /*::]*/: '#REF!',
    /*::[*/0x1D /*::]*/: '#NAME?',
    /*::[*/0x24 /*::]*/: '#NUM!',
    /*::[*/0x2A /*::]*/: '#N/A',
    /*::[*/0x2B /*::]*/: '#GETTING_DATA',
    /*::[*/0xFF /*::]*/: '#WTF?',
}

export const RBErr = evert_num(BErr)

/* [MS-XLSB] 2.4.321 BrtColor */
export function parse_BrtColor(data, length: number) {
    const out = {}
    const d = data.read_shift(1)
    out.fValidRGB = d & 1
    out.xColorType = d >>> 1
    out.index = data.read_shift(1)
    out.nTintAndShade = data.read_shift(2, 'i')
    out.bRed = data.read_shift(1)
    out.bGreen = data.read_shift(1)
    out.bBlue = data.read_shift(1)
    out.bAlpha = data.read_shift(1)
}

/* [MS-XLSB] 2.5.52 */
export function parse_FontFlags(data, length: number) {
    const d = data.read_shift(1)
    data.l++
    const out = {
        fItalic: d & 0x2,
        fStrikeout: d & 0x8,
        fOutline: d & 0x10,
        fShadow: d & 0x20,
        fCondense: d & 0x40,
        fExtend: d & 0x80,
    }
    return out
}
