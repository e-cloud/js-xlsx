import * as cptable from 'codepage/dist/cpexcel.full.js'
import * as fs from 'fs'
import * as JSZip from 'jszip'
import * as Base64 from './04_base64'
import { has_buf, s2a } from './05_buf'
import * as CFB from './18_cfb'
import { DBF, DIF, PRN, SYLK } from './40_harb'
import { WK_ } from './41_lotus'
import { _ssfopts, resetSSFOpts } from './66_wscommon'
import { parse_xlml } from './75_xlml'
import { parse_xlscfb } from './76_xls'
import { parse_xlsxcfb, parse_zip } from './85_parsezip'

function firstbyte(f: RawData, o ?: TypeOpts): Array<number> {
    let x = ''
    switch ((o || {}).type || 'base64') {
        case 'buffer':
            return [f[0], f[1], f[2], f[3]]
        case 'base64':
            x = Base64.decode(f.substr(0, 24))
            break
        case 'binary':
            x = f
            break
        case 'array':
            return [f[0], f[1], f[2], f[3]]
        default:
            throw new Error(`Unrecognized type ${o ? o.type : 'undefined'}`)
    }
    return [x.charCodeAt(0), x.charCodeAt(1), x.charCodeAt(2), x.charCodeAt(3)]
}

export function read_cfb(cfb, opts ?: ParseOpts): Workbook {
    if (cfb.find('EncryptedPackage')) {
        return parse_xlsxcfb(cfb, opts)
    }
    return parse_xlscfb(cfb, opts)
}

export function read_zip(data: RawData, opts ?: ParseOpts): Workbook {
    /*:: if(!jszip) throw new Error("JSZip is not available"); */
    let zip

    const d = data
    const o = opts || {}
    if (!o.type) {
        o.type = has_buf && Buffer.isBuffer(data) ? 'buffer' : 'base64'
    }
    switch (o.type) {
        case 'base64':
            zip = new JSZip(d, { base64: true })
            break
        case 'binary':
        case 'array':
            zip = new JSZip(d, { base64: false })
            break
        case 'buffer':
            zip = new JSZip(d)
            break
        default:
            throw new Error(`Unrecognized type ${o.type}`)
    }
    return parse_zip(zip, o)
}

function read_utf16(data: RawData, o: ParseOpts): Workbook {
    let d = data
    if (o.type == 'base64') {
        d = Base64.decode(d)
    }
    d = cptable.utils.decode(1200, d.slice(2))
    o.type = 'binary'
    if (d.charCodeAt(0) == 0x3C) {
        return parse_xlml(d, o)
    }
    return PRN.to_workbook(d, o)
}

export function readSync(data: RawData, opts ?: ParseOpts): Workbook {
    let zip
    let d = data
    let n = [0]
    const o = opts || {}
    resetSSFOpts({})
    if (o.dateNF) {
        _ssfopts.dateNF = o.dateNF
    }
    if (!o.type) {
        o.type = has_buf && Buffer.isBuffer(data) ? 'buffer' : 'base64'
    }
    if (o.type == 'file') {
        o.type = 'buffer'
        d = fs.readFileSync(data)
    }
    switch ((n = firstbyte(d, o))[0]) {
        case 0xD0:
            return read_cfb(CFB.read(d, o), o)
        case 0x09:
            return parse_xlscfb(s2a(o.type === 'base64' ? Base64.decode(d) : d), o)
        case 0x3C:
            return parse_xlml(d, o)
        case 0x49:
            if (n[1] == 0x44) {
                return SYLK.to_workbook(d, o)
            }
            break
        case 0x54:
            if (n[1] == 0x41 && n[2] == 0x42 && n[3] == 0x4C) {
                return DIF.to_workbook(d, o)
            }
            break
        case 0x50:
            if (n[1] == 0x4B && n[2] < 0x20 && n[3] < 0x20) {
                return read_zip(d, o)
            }
            break
        case 0xEF:
            return parse_xlml(d, o)
        case 0xFF:
            if (n[1] == 0xFE) {
                return read_utf16(d, o)
            }
            break
        case 0x00:
            if (n[1] == 0x00 && n[2] >= 0x02 && n[3] == 0x00) {
                return WK_.to_workbook(d, o)
            }
            break
        case 0x03:
        case 0x83:
        case 0x8B:
            return DBF.to_workbook(d, o)
    }
    if (n[2] <= 12 && n[3] <= 31) {
        return DBF.to_workbook(d, o)
    }
    if (0x20 > n[0] || n[0] > 0x7F) {
        throw new Error(`Unsupported file ${n.join('|')}`)
    }
    return PRN.to_workbook(d, o)
}

export function readFileSync(filename: string, opts ?: ParseOpts): Workbook {
    const o = opts || {}
    o.type = 'file'
    return readSync(filename, o)
}
