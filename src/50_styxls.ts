import { parsenoop } from './23_binutils'
import { parse_LongRGBA } from './38_xlstypes'

/* [MS-XLS] 2.4.326 TODO: payload is a zip file */
export function parse_Theme(blob, length, opts) {
    const dwThemeVersion = blob.read_shift(4)
    if (dwThemeVersion === 124226) {
        return
    }
    blob.l += length - 4
}

/* 2.5.49 */
export function parse_ColorTheme(blob, length) {
    return blob.read_shift(4)
}

/* 2.5.155 */
export function parse_FullColorExt(blob, length) {
    const o = {}
    o.xclrType = blob.read_shift(2)
    o.nTintShade = blob.read_shift(2)
    switch (o.xclrType) {
        case 0:
            blob.l += 4
            break
        case 1:
            o.xclrValue = parse_IcvXF(blob, 4)
            break
        case 2:
            o.xclrValue = parse_LongRGBA(blob, 4)
            break
        case 3:
            o.xclrValue = parse_ColorTheme(blob, 4)
            break
        case 4:
            blob.l += 4
            break
    }
    blob.l += 8
    return o
}

/* 2.5.164 TODO: read 7 bits*/
export function parse_IcvXF(blob, length) {
    return parsenoop(blob, length)
}

/* 2.5.280 */
export function parse_XFExtGradient(blob, length) {
    return parsenoop(blob, length)
}

/* 2.5.108 */
export function parse_ExtProp(blob, length) {
    const extType = blob.read_shift(2)
    const cb = blob.read_shift(2)
    const o = [extType]
    switch (extType) {
        case 0x04:
        case 0x05:
        case 0x07:
        case 0x08:
        case 0x09:
        case 0x0A:
        case 0x0B:
        case 0x0D:
            o[1] = parse_FullColorExt(blob, cb)
            break
        case 0x06:
            o[1] = parse_XFExtGradient(blob, cb)
            break
        case 0x0E:
        case 0x0F:
            o[1] = blob.read_shift(cb === 5 ? 1 : 2)
            break
        default:
            throw new Error(`Unrecognized ExtProp type: ${extType} ${cb}`)
    }
    return o
}

/* 2.4.355 */
export function parse_XFExt(blob, length) {
    const end = blob.l + length
    blob.l += 2
    const ixfe = blob.read_shift(2)
    blob.l += 2
    let cexts = blob.read_shift(2)
    const ext = []
    while (cexts-- > 0) {
        ext.push(parse_ExtProp(blob, end - blob.l))
    }
    return { ixfe, ext }
}

/* xf is an XF, see parse_XFExt for xfext */
export function update_xfext(xf, xfext) {
    xfext.forEach(function (xfe) {
        switch (xfe[0]) {/* 2.5.108 extPropData */
            case 0x04:
                break /* foreground color */
            case 0x05:
                break /* background color */
            case 0x06:
                break /* gradient fill */
            case 0x07:
                break /* top cell border color */
            case 0x08:
                break /* bottom cell border color */
            case 0x09:
                break /* left cell border color */
            case 0x0a:
                break /* right cell border color */
            case 0x0b:
                break /* diagonal cell border color */
            case 0x0d:
                break /* text color */
            case 0x0e:
                break /* font scheme */
            case 0x0f:
                break /* indentation level */
        }
    })
}
