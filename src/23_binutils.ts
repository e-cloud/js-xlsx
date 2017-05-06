import cptable from 'codepage/dist/cpexcel.full.js'
import { _getchar, current_codepage } from './02_codepage'
import * as buf from './05_buf'
import { has_buf, new_raw_buf } from './05_buf'

export function read_double_le(b, idx: number): number {
    const s = 1 - 2 * (b[idx + 7] >>> 7)
    let e = ((b[idx + 7] & 0x7f) << 4) + (b[idx + 6] >>> 4 & 0x0f)
    let m = b[idx + 6] & 0x0f
    for (let i = 5; i >= 0; --i) {
        m = m * 256 + b[idx + i]
    }
    if (e == 0x7ff) {
        return m == 0 ? s * Infinity : NaN
    }
    if (e == 0) {
        e = -1022
    } else {
        e -= 1023
        m += Math.pow(2, 52)
    }
    return s * Math.pow(2, e - 52) * m
}

export function write_double_le(b, v: number, idx: number) {
    const bs = (v < 0 || 1 / v == -Infinity ? 1 : 0) << 7
    let e = 0
    let m = 0
    const av = bs ? -v : v
    if (!isFinite(av)) {
        e = 0x7ff
        m = isNaN(v) ? 0x6969 : 0
    } else {
        e = Math.floor(Math.log(av) * Math.LOG2E)
        m = v * Math.pow(2, 52 - e)
        if (e <= -1023 && (!isFinite(m) || m < Math.pow(2, 52))) {
            e = -1022
        } else {
            m -= Math.pow(2, 52)
            e += 1023
        }
    }
    for (let i = 0; i <= 5; ++i, m /= 256) {
        b[idx + i] = m & 0xff
    }
    b[idx + 6] = (e & 0x0f) << 4 | m & 0xf
    b[idx + 7] = e >> 4 | bs
}

export let __toBuffer
let ___toBuffer
__toBuffer = ___toBuffer = function toBuffer_(bufs) {
    const x = []
    for (let i = 0; i < bufs[0].length; ++i) {
        x.push(...bufs[0][i])
    }
    return x
}
export let __utf16le
let ___utf16le
__utf16le = ___utf16le = function utf16le_(b, s, e) {
    const ss = []
    for (let i = s; i < e; i += 2) {
        ss.push(String.fromCharCode(__readUInt16LE(b, i)))
    }
    return ss.join('')
}
export let __hexlify
let ___hexlify
__hexlify = ___hexlify = function hexlify_(b, s, l) {
    return b.slice(s, s + l).map(function (x) {
        return (x < 16 ? '0' : '') + x.toString(16)
    }).join('')
}
export let __utf8
let ___utf8
__utf8 = ___utf8 = function (b, s, e) {
    const ss = []
    for (let i = s; i < e; i++) {
        ss.push(String.fromCharCode(__readUInt8(b, i)))
    }
    return ss.join('')
}
export let __lpstr
let ___lpstr
__lpstr = ___lpstr = function lpstr_(b, i) {
    const len = __readUInt32LE(b, i)
    return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : ''
}
export let __lpwstr
let ___lpwstr
__lpwstr = ___lpwstr = function lpwstr_(b, i) {
    const len = 2 * __readUInt32LE(b, i)
    return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : ''
}
export let __lpp4
let ___lpp4
__lpp4 = ___lpp4 = function lpp4_(b, i) {
    const len = __readUInt32LE(b, i)
    return len > 0 ? __utf16le(b, i + 4, i + 4 + len) : ''
}
export let __8lpp4
let ___8lpp4
__8lpp4 = ___8lpp4 = function lpp4_8(b, i) {
    const len = __readUInt32LE(b, i)
    return len > 0 ? __utf8(b, i + 4, i + 4 + len) : ''
}
export let __double
let ___double
__double = ___double = function (b, idx) {
    return read_double_le(b, idx)
}

export let is_buf = function is_buf_a(a) {
    return Array.isArray(a)
}

if (has_buf /*:: && typeof Buffer != 'undefined'*/) {
    __utf16le = function utf16le_b(b, s, e) {
        if (!Buffer.isBuffer(b)) {
            return ___utf16le(b, s, e)
        }
        return b.toString('utf16le', s, e)
    }
    __hexlify = function (b, s, l) {
        return Buffer.isBuffer(b) ? b.toString('hex', s, s + l) : ___hexlify(b, s, l)
    }
    __lpstr = function lpstr_b(b, i) {
        if (!Buffer.isBuffer(b)) {
            return ___lpstr(b, i)
        }
        const len = b.readUInt32LE(i)
        return len > 0 ? b.toString('utf8', i + 4, i + 4 + len - 1) : ''
    }
    __lpwstr = function lpwstr_b(b, i) {
        if (!Buffer.isBuffer(b)) {
            return ___lpwstr(b, i)
        }
        const len = 2 * b.readUInt32LE(i)
        return b.toString('utf16le', i + 4, i + 4 + len - 1)
    }
    __lpp4 = function lpp4_b(b, i) {
        if (!Buffer.isBuffer(b)) {
            return ___lpp4(b, i)
        }
        const len = b.readUInt32LE(i)
        return b.toString('utf16le', i + 4, i + 4 + len)
    }
    __8lpp4 = function lpp4_8b(b, i) {
        if (!Buffer.isBuffer(b)) {
            return ___8lpp4(b, i)
        }
        const len = b.readUInt32LE(i)
        return b.toString('utf8', i + 4, i + 4 + len)
    }
    __utf8 = function utf8_b(b, s, e) {
        return b.toString('utf8', s, e)
    }
    __toBuffer = function (bufs) {
        return bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0]) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs)
    }
    buf.bconcat = function (bufs) {
        return Buffer.isBuffer(bufs[0]) ? Buffer.concat(bufs) : [].concat(...bufs)
    }
    __double = function double_(b, i) {
        if (Buffer.isBuffer(b) /*::&& b instanceof Buffer*/) {
            return b.readDoubleLE(i)
        }
        return ___double(b, i)
    }
    is_buf = function is_buf_b(a) {
        return Buffer.isBuffer(a) || Array.isArray(a)
    }
}

/* from js-xls */
if (typeof cptable !== 'undefined') {
    __utf16le = function (b, s, e) {
        return cptable.utils.decode(1200, b.slice(s, e))
    }
    __utf8 = function (b, s, e) {
        return cptable.utils.decode(65001, b.slice(s, e))
    }
    __lpstr = function (b, i) {
        const len = __readUInt32LE(b, i)
        return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i + 4, i + 4 + len - 1)) : ''
    }
    __lpwstr = function (b, i) {
        const len = 2 * __readUInt32LE(b, i)
        return len > 0 ? cptable.utils.decode(1200, b.slice(i + 4, i + 4 + len - 1)) : ''
    }
    __lpp4 = function (b, i) {
        const len = __readUInt32LE(b, i)
        return len > 0 ? cptable.utils.decode(1200, b.slice(i + 4, i + 4 + len)) : ''
    }
    __8lpp4 = function (b, i) {
        const len = __readUInt32LE(b, i)
        return len > 0 ? cptable.utils.decode(65001, b.slice(i + 4, i + 4 + len)) : ''
    }
}

const __readUInt8 = function (b, idx) {
    return b[idx]
}
export const __readUInt16LE = function (b, idx) {
    return b[idx + 1] * (1 << 8) + b[idx]
}
const __readInt16LE = function (b, idx) {
    const u = b[idx + 1] * (1 << 8) + b[idx]
    return u < 0x8000 ? u : (0xffff - u + 1) * -1
}
export const __readUInt32LE = function (b, idx) {
    return b[idx + 3] * (1 << 24) + (b[idx + 2] << 16) + (b[idx + 1] << 8) + b[idx]
}
export const __readInt32LE = function (b, idx) {
    return b[idx + 3] << 24 | b[idx + 2] << 16 | b[idx + 1] << 8 | b[idx]
}

const ___unhexlify = function (s) {
    return s.match(/../g).map(function (x) {
        return parseInt(x, 16)
    })
}
const __unhexlify = typeof Buffer !== 'undefined' ? function (s) {
    return Buffer.isBuffer(s) ? new Buffer(s, 'hex') : ___unhexlify(s)
} : ___unhexlify

export function ReadShift(size: number, t ?: string) {
    let o = ''
    let oI
    let oR
    const oo = []
    let w
    let vv
    let i
    let loc
    switch (t) {
        case 'dbcs':
            loc = this.l
            if (has_buf && Buffer.isBuffer(this)) {
                o = this.slice(this.l, this.l + 2 * size).toString('utf16le')
            } else {
                for (i = 0; i != size; ++i) {
                    o += String.fromCharCode(__readUInt16LE(this, loc))
                    loc += 2
                }
            }
            size *= 2
            break

        case 'utf8':
            o = __utf8(this, this.l, this.l + size)
            break
        case 'utf16le':
            size *= 2
            o = __utf16le(this, this.l, this.l + size)
            break

        case 'wstr':
            if (typeof cptable !== 'undefined') {
                o = cptable.utils.decode(current_codepage, this.slice(this.l, this.l + 2 * size))
            } else {
                return ReadShift.call(this, size, 'dbcs')
            }
            size = 2 * size
            break

        /* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
        case 'lpstr':
            o = __lpstr(this, this.l)
            size = 5 + o.length
            break
        /* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
        case 'lpwstr':
            o = __lpwstr(this, this.l)
            size = 5 + o.length
            if (o[o.length - 1] == '\0') {
                size += 2
            }
            break
        /* [MS-OFFCRYPTO] 2.1.2 Length-Prefixed Padded Unicode String (UNICODE-LP-P4) */
        case 'lpp4':
            size = 4 + __readUInt32LE(this, this.l)
            o = __lpp4(this, this.l)
            if (size & 0x02) {
                size += 2
            }
            break
        /* [MS-OFFCRYPTO] 2.1.3 Length-Prefixed UTF-8 String (UTF-8-LP-P4) */
        case '8lpp4':
            size = 4 + __readUInt32LE(this, this.l)
            o = __8lpp4(this, this.l)
            if (size & 0x03) {
                size += 4 - (size & 0x03)
            }
            break

        case 'cstr':
            size = 0
            o = ''
            while ((w = __readUInt8(this, this.l + size++)) !== 0) {
                oo.push(_getchar(w))
            }
            o = oo.join('')
            break
        case '_wstr':
            size = 0
            o = ''
            while ((w = __readUInt16LE(this, this.l + size)) !== 0) {
                oo.push(_getchar(w))
                size += 2
            }
            size += 2
            o = oo.join('')
            break

        /* sbcs and dbcs support continue records in the SST way TODO codepages */
        case 'dbcs-cont':
            o = ''
            loc = this.l
            for (i = 0; i != size; ++i) {
                if (this.lens && this.lens.includes(loc)) {
                    w = __readUInt8(this, loc)
                    this.l = loc + 1
                    vv = ReadShift.call(this, size - i, w ? 'dbcs-cont' : 'sbcs-cont')
                    return oo.join('') + vv
                }
                oo.push(_getchar(__readUInt16LE(this, loc)))
                loc += 2
            }
            o = oo.join('')
            size *= 2
            break

        case 'sbcs-cont':
            o = ''
            loc = this.l
            for (i = 0; i != size; ++i) {
                if (this.lens && this.lens.includes(loc)) {
                    w = __readUInt8(this, loc)
                    this.l = loc + 1
                    vv = ReadShift.call(this, size - i, w ? 'dbcs-cont' : 'sbcs-cont')
                    return oo.join('') + vv
                }
                oo.push(_getchar(__readUInt8(this, loc)))
                loc += 1
            }
            o = oo.join('')
            break

        default:
            switch (size) {
                case 1:
                    oI = __readUInt8(this, this.l)
                    this.l++
                    return oI
                case 2:
                    oI = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l)
                    this.l += 2
                    return oI
                case 4:
                    if (t === 'i' || (this[this.l + 3] & 0x80) === 0) {
                        oI = __readInt32LE(this, this.l)
                        this.l += 4
                        return oI
                    } else {
                        oR = __readUInt32LE(this, this.l)
                        this.l += 4
                    }
                    return oR
                case 8:
                    if (t === 'f') {
                        oR = __double(this, this.l)
                        this.l += 8
                        return oR
                    }
                /* falls through */
                case 16:
                    o = __hexlify(this, this.l, size)
                    break
            }
    }
    this.l += size
    return o
}

const __writeUInt16LE = function (b, val, idx) {
    b[idx] = val & 0xFF
    b[idx + 1] = val >>> 8 & 0xFF
}
const __writeUInt32LE = function (b, val, idx) {
    b[idx] = val & 0xFF
    b[idx + 1] = val >>> 8 & 0xFF
    b[idx + 2] = val >>> 16 & 0xFF
    b[idx + 3] = val >>> 24 & 0xFF
}
const __writeInt32LE = function (b, val, idx) {
    b[idx] = val & 0xFF
    b[idx + 1] = val >> 8 & 0xFF
    b[idx + 2] = val >> 16 & 0xFF
    b[idx + 3] = val >> 24 & 0xFF
}

export function WriteShift(t: number, val: string | number, f?: string) {
    let size = 0
    let i = 0
    if (f === 'dbcs') {
        /*:: if(typeof val !== 'string') throw new Error("unreachable"); */
        for (i = 0; i != val.length; ++i) {
            __writeUInt16LE(this, val.charCodeAt(i), this.l + 2 * i)
        }
        size = 2 * val.length
    } else if (f === 'sbcs') {
        /*:: if(typeof val !== 'string') throw new Error("unreachable"); */
        for (i = 0; i != val.length; ++i) {
            this[this.l + i] = val.charCodeAt(i) & 0xFF
        }
        size = val.length
    } else /*:: if(typeof val === 'number') */{
        switch (t) {
            case 1:
                size = 1
                this[this.l] = val & 0xFF
                break
            case 2:
                size = 2
                this[this.l] = val & 0xFF
                val >>>= 8
                this[this.l + 1] = val & 0xFF
                break
            case 3:
                size = 3
                this[this.l] = val & 0xFF
                val >>>= 8
                this[this.l + 1] = val & 0xFF
                val >>>= 8
                this[this.l + 2] = val & 0xFF
                break
            case 4:
                size = 4
                __writeUInt32LE(this, val, this.l)
                break
            case 8:
                size = 8
                if (f === 'f') {
                    write_double_le(this, val, this.l)
                    break
                }
            /* falls through */
            case 16:
                break
            case -4:
                size = 4
                __writeInt32LE(this, val, this.l)
                break
        }
    }
    this.l += size
    return this
}

export function CheckField(hexstr, fld) {
    const m = __hexlify(this, this.l, hexstr.length >> 1)
    if (m !== hexstr) {
        throw `${fld}Expected ${hexstr} saw ${m}`
    }
    this.l += hexstr.length >> 1
}

export function prep_blob(blob, pos: number) {
    blob.l = pos
    blob.read_shift = ReadShift
    blob.chk = CheckField
    blob.write_shift = WriteShift
}

export function parsenoop(blob, length: number) {
    blob.l += length
}

export function parsenooplog(blob, length: number) {
    if (typeof console != 'undefined') {
        console.log(blob.slice(blob.l, blob.l + length))
    }
    blob.l += length
}

export function writenoop(blob, length: number) {
    blob.l += length
}

export function new_buf(sz: number): Block {
    const o = new_raw_buf(sz)
    prep_blob(o, 0)
    return o
}
