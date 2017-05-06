import cptable from 'codepage/dist/cpexcel.full.js'
import { new_raw_buf } from './05_buf'
import { parseuint16 } from './38_xlstypes'

function _JS2ANSI(str: string): Array<number> {
    if (typeof cptable !== 'undefined') {
        return cptable.utils.encode(1252, str)
    }
    const o = []
    const oo = str.split('')
    for (let i = 0; i < oo.length; ++i) {
        o[i] = oo[i].charCodeAt(0)
    }
    return o
}

/* [MS-OFFCRYPTO] 2.1.4 Version */
function parse_CRYPTOVersion(blob, length?: number) {
    const o = {}
    o.Major = blob.read_shift(2)
    o.Minor = blob.read_shift(2)
    return o
}

/* [MS-OFFCRYPTO] 2.1.5 DataSpaceVersionInfo */
export function parse_DataSpaceVersionInfo(blob, length?) {
    const o = {}
    o.id = blob.read_shift(0, 'lpp4')
    o.R = parse_CRYPTOVersion(blob, 4)
    o.U = parse_CRYPTOVersion(blob, 4)
    o.W = parse_CRYPTOVersion(blob, 4)
    return o
}

/* [MS-OFFCRYPTO] 2.1.6.1 DataSpaceMapEntry Structure */
export function parse_DataSpaceMapEntry(blob) {
    const len = blob.read_shift(4)
    const end = blob.l + len - 4
    const o = {}
    let cnt = blob.read_shift(4)
    const comps = []
    while (cnt-- > 0) {
        /* [MS-OFFCRYPTO] 2.1.6.2 DataSpaceReferenceComponent Structure */
        const rc = {}
        rc.t = blob.read_shift(4)
        rc.v = blob.read_shift(0, 'lpp4')
        comps.push(rc)
    }
    o.name = blob.read_shift(0, 'lpp4')
    o.comps = comps
    return o
}

/* [MS-OFFCRYPTO] 2.1.6 DataSpaceMap */
export function parse_DataSpaceMap(blob, length?) {
    const o = []
    blob.l += 4 // must be 0x8
    let cnt = blob.read_shift(4)
    while (cnt-- > 0) {
        o.push(parse_DataSpaceMapEntry(blob))
    }
    return o
}

/* [MS-OFFCRYPTO] 2.1.7 DataSpaceDefinition */
export function parse_DataSpaceDefinition(blob, length?) {
    const o = []
    blob.l += 4 // must be 0x8
    let cnt = blob.read_shift(4)
    while (cnt-- > 0) {
        o.push(blob.read_shift(0, 'lpp4'))
    }
    return o
}

/* [MS-OFFCRYPTO] 2.1.8 DataSpaceDefinition */
export function parse_TransformInfoHeader(blob, length?) {
    const o = {}
    const len = blob.read_shift(4)
    const tgt = blob.l + len - 4
    blob.l += 4 // must be 0x1
    o.id = blob.read_shift(0, 'lpp4')
    // tgt == len
    o.name = blob.read_shift(0, 'lpp4')
    o.R = parse_CRYPTOVersion(blob, 4)
    o.U = parse_CRYPTOVersion(blob, 4)
    o.W = parse_CRYPTOVersion(blob, 4)
    return o
}

export function parse_Primary(blob, length?) {
    /* [MS-OFFCRYPTO] 2.2.6 IRMDSTransformInfo */
    const hdr = parse_TransformInfoHeader(blob)
    /* [MS-OFFCRYPTO] 2.1.9 EncryptionTransformInfo */
    hdr.ename = blob.read_shift(0, '8lpp4')
    hdr.blksz = blob.read_shift(4)
    hdr.cmode = blob.read_shift(4)
    if (blob.read_shift(4) != 0x04) {
        throw new Error('Bad !Primary record')
    }
    return hdr
}

/* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
export function parse_EncryptionHeader(blob, length: number) {
    const tgt = blob.l + length
    const o = {}
    o.Flags = blob.read_shift(4) & 0x3F
    blob.l += 4
    o.AlgID = blob.read_shift(4)
    let valid = false
    switch (o.AlgID) {
        case 0x660E:
        case 0x660F:
        case 0x6610:
            valid = o.Flags == 0x24
            break
        case 0x6801:
            valid = o.Flags == 0x04
            break
        case 0:
            valid = o.Flags == 0x10 || o.Flags == 0x04 || o.Flags == 0x24
            break
        default:
            throw `Unrecognized encryption algorithm: ${o.AlgID}`
    }
    if (!valid) {
        throw new Error('Encryption Flags/AlgID mismatch')
    }
    o.AlgIDHash = blob.read_shift(4)
    o.KeySize = blob.read_shift(4)
    o.ProviderType = blob.read_shift(4)
    blob.l += 8
    o.CSPName = blob.read_shift(tgt - blob.l >> 1, 'utf16le').slice(0, -1)
    blob.l = tgt
    return o
}

/* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
export function parse_EncryptionVerifier(blob, length: number) {
    const o = {}
    blob.l += 4 // SaltSize must be 0x10
    o.Salt = blob.slice(blob.l, blob.l + 16)
    blob.l += 16
    o.Verifier = blob.slice(blob.l, blob.l + 16)
    blob.l += 16
    const sz = blob.read_shift(4)
    o.VerifierHash = blob.slice(blob.l, blob.l + sz)
    blob.l += sz
    return o
}

/* [MS-OFFCRYPTO] 2.3.4.* EncryptionInfo Stream */
export function parse_EncryptionInfo(blob, length?) {
    const vers = parse_CRYPTOVersion(blob)
    switch (vers.Minor) {
        case 0x02:
            return parse_EncInfoStd(blob, vers)
        case 0x03:
            return parse_EncInfoExt(blob, vers)
        case 0x04:
            return parse_EncInfoAgl(blob, vers)
    }
    throw new Error(`ECMA-376 Encryped file unrecognized Version: ${vers.Minor}`)
}

/* [MS-OFFCRYPTO] 2.3.4.5  EncryptionInfo Stream (Standard Encryption) */
function parse_EncInfoStd(blob, vers) {
    const flags = blob.read_shift(4)
    if ((flags & 0x3F) != 0x24) {
        throw new Error('EncryptionInfo mismatch')
    }
    const sz = blob.read_shift(4)
    const tgt = blob.l + sz
    const hdr = parse_EncryptionHeader(blob, sz)
    const verifier = parse_EncryptionVerifier(blob, blob.length - blob.l)
    return { t: 'Std', h: hdr, v: verifier }
}

/* [MS-OFFCRYPTO] 2.3.4.6  EncryptionInfo Stream (Extensible Encryption) */
function parse_EncInfoExt(blob, vers) {
    throw new Error('File is password-protected: ECMA-376 Extensible')
}

/* [MS-OFFCRYPTO] 2.3.4.10 EncryptionInfo Stream (Agile Encryption) */
function parse_EncInfoAgl(blob, vers) {
    throw new Error('File is password-protected: ECMA-376 Agile')
}

/* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
function parse_RC4CryptoHeader(blob, length: number) {
    const o = {}
    const vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4)
    length -= 4
    if (vers.Minor != 2) {
        throw `unrecognized minor version code: ${vers.Minor}`
    }
    if (vers.Major > 4 || vers.Major < 2) {
        throw `unrecognized major version code: ${vers.Major}`
    }
    o.Flags = blob.read_shift(4)
    length -= 4
    const sz = blob.read_shift(4)
    length -= 4
    o.EncryptionHeader = parse_EncryptionHeader(blob, sz)
    length -= sz
    o.EncryptionVerifier = parse_EncryptionVerifier(blob, length)
    return o
}

/* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
function parse_RC4Header(blob, length: number) {
    const o = {}
    const vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4)
    length -= 4
    if (vers.Major != 1 || vers.Minor != 1) {
        throw `unrecognized version code ${vers.Major} : ${vers.Minor}`
    }
    o.Salt = blob.read_shift(16)
    o.EncryptedVerifier = blob.read_shift(16)
    o.EncryptedVerifierHash = blob.read_shift(16)
    return o
}

/* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
export function crypto_CreatePasswordVerifier_Method1(password: string) {
    let Verifier = 0x0000
    let PasswordArray
    const PasswordDecoded = _JS2ANSI(password)
    const len = PasswordDecoded.length + 1
    let i
    let PasswordByte
    let Intermediate1
    let Intermediate2
    let Intermediate3
    PasswordArray = new_raw_buf(len)
    PasswordArray[0] = PasswordDecoded.length
    for (i = 1; i != len; ++i) {
        PasswordArray[i] = PasswordDecoded[i - 1]
    }
    for (i = len - 1; i >= 0; --i) {
        PasswordByte = PasswordArray[i]
        Intermediate1 = (Verifier & 0x4000) === 0x0000 ? 0 : 1
        Intermediate2 = Verifier << 1 & 0x7FFF
        Intermediate3 = Intermediate1 | Intermediate2
        Verifier = Intermediate3 ^ PasswordByte
    }
    return Verifier ^ 0xCE4B
}

/* [MS-OFFCRYPTO] 2.3.7.2 Binary Document XOR Array Initialization */
const crypto_CreateXorArray_Method1 = function () {
    const PadArray = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00]
    const InitialCode = [
        0xE1F0,
        0x1D0F,
        0xCC9C,
        0x84C0,
        0x110C,
        0x0E10,
        0xF1CE,
        0x313E,
        0x1872,
        0xE139,
        0xD40F,
        0x84F9,
        0x280C,
        0xA96A,
        0x4EC3,
    ]
    const XorMatrix = [
        0xAEFC,
        0x4DD9,
        0x9BB2,
        0x2745,
        0x4E8A,
        0x9D14,
        0x2A09,
        0x7B61,
        0xF6C2,
        0xFDA5,
        0xEB6B,
        0xC6F7,
        0x9DCF,
        0x2BBF,
        0x4563,
        0x8AC6,
        0x05AD,
        0x0B5A,
        0x16B4,
        0x2D68,
        0x5AD0,
        0x0375,
        0x06EA,
        0x0DD4,
        0x1BA8,
        0x3750,
        0x6EA0,
        0xDD40,
        0xD849,
        0xA0B3,
        0x5147,
        0xA28E,
        0x553D,
        0xAA7A,
        0x44D5,
        0x6F45,
        0xDE8A,
        0xAD35,
        0x4A4B,
        0x9496,
        0x390D,
        0x721A,
        0xEB23,
        0xC667,
        0x9CEF,
        0x29FF,
        0x53FE,
        0xA7FC,
        0x5FD9,
        0x47D3,
        0x8FA6,
        0x0F6D,
        0x1EDA,
        0x3DB4,
        0x7B68,
        0xF6D0,
        0xB861,
        0x60E3,
        0xC1C6,
        0x93AD,
        0x377B,
        0x6EF6,
        0xDDEC,
        0x45A0,
        0x8B40,
        0x06A1,
        0x0D42,
        0x1A84,
        0x3508,
        0x6A10,
        0xAA51,
        0x4483,
        0x8906,
        0x022D,
        0x045A,
        0x08B4,
        0x1168,
        0x76B4,
        0xED68,
        0xCAF1,
        0x85C3,
        0x1BA7,
        0x374E,
        0x6E9C,
        0x3730,
        0x6E60,
        0xDCC0,
        0xA9A1,
        0x4363,
        0x86C6,
        0x1DAD,
        0x3331,
        0x6662,
        0xCCC4,
        0x89A9,
        0x0373,
        0x06E6,
        0x0DCC,
        0x1021,
        0x2042,
        0x4084,
        0x8108,
        0x1231,
        0x2462,
        0x48C4,
    ]
    const Ror = function (Byte) {
        return (Byte / 2 | Byte * 128) & 0xFF
    }
    const XorRor = function (byte1, byte2) {
        return Ror(byte1 ^ byte2)
    }
    const CreateXorKey_Method1 = function (Password) {
        let XorKey = InitialCode[Password.length - 1]
        let CurrentElement = 0x68
        for (let i = Password.length - 1; i >= 0; --i) {
            let Char = Password[i]
            for (let j = 0; j != 7; ++j) {
                if (Char & 0x40) {
                    XorKey ^= XorMatrix[CurrentElement]
                }
                Char *= 2
                --CurrentElement
            }
        }
        return XorKey
    }
    return function (password: string) {
        const Password = _JS2ANSI(password)
        const XorKey = CreateXorKey_Method1(Password)
        let Index = Password.length
        const ObfuscationArray = new_raw_buf(16)
        for (let i = 0; i != 16; ++i) {
            ObfuscationArray[i] = 0x00
        }
        let Temp
        let PasswordLastChar
        let PadIndex
        if ((Index & 1) === 1) {
            Temp = XorKey >> 8
            ObfuscationArray[Index] = XorRor(PadArray[0], Temp)
            --Index
            Temp = XorKey & 0xFF
            PasswordLastChar = Password[Password.length - 1]
            ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp)
        }
        while (Index > 0) {
            --Index
            Temp = XorKey >> 8
            ObfuscationArray[Index] = XorRor(Password[Index], Temp)
            --Index
            Temp = XorKey & 0xFF
            ObfuscationArray[Index] = XorRor(Password[Index], Temp)
        }
        Index = 15
        PadIndex = 15 - Password.length
        while (PadIndex > 0) {
            Temp = XorKey >> 8
            ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp)
            --Index
            --PadIndex
            Temp = XorKey & 0xFF
            ObfuscationArray[Index] = XorRor(Password[Index], Temp)
            --Index
            --PadIndex
        }
        return ObfuscationArray
    }
}()

/* [MS-OFFCRYPTO] 2.3.7.3 Binary Document XOR Data Transformation Method 1 */
const crypto_DecryptData_Method1 = function (password: string, Data, XorArrayIndex, XorArray, O?) {
    /* If XorArray is set, use it; if O is not set, make changes in-place */
    if (!O) {
        O = Data
    }
    if (!XorArray) {
        XorArray = crypto_CreateXorArray_Method1(password)
    }
    let Index
    let Value
    for (Index = 0; Index != Data.length; ++Index) {
        Value = Data[Index]
        Value ^= XorArray[XorArrayIndex]
        Value = (Value >> 5 | Value << 3) & 0xFF
        O[Index] = Value
        ++XorArrayIndex
    }
    return [O, XorArrayIndex, XorArray]
}

const crypto_MakeXorDecryptor = function (password: string) {
    let XorArrayIndex = 0
    const XorArray = crypto_CreateXorArray_Method1(password)
    return function (Data) {
        const O = crypto_DecryptData_Method1('', Data, XorArrayIndex, XorArray)
        XorArrayIndex = O[1]
        return O[0]
    }
}

/* 2.5.343 */
export function parse_XORObfuscation(blob, length, opts, out) {
    const o = { key: parseuint16(blob), verificationBytes: parseuint16(blob) }

    if (opts.password) {
        o.verifier = crypto_CreatePasswordVerifier_Method1(opts.password)
    }
    out.valid = o.verificationBytes === o.verifier
    if (out.valid) {
        out.insitu_decrypt = crypto_MakeXorDecryptor(opts.password)
    }
    return o
}

/* 2.4.117 */
export function parse_FilePassHeader(blob, length: number, oo) {
    const o = oo || {}
    o.Info = blob.read_shift(2)
    blob.l -= 2
    if (o.Info === 1) {
        o.Data = parse_RC4Header(blob, length)
    } else {
        o.Data = parse_RC4CryptoHeader(blob, length)
    }
    return o
}

export function parse_FilePass(blob, length: number, opts) {
    const o = { Type: blob.read_shift(2) }
    /* wEncryptionType */
    if (o.Type) {
        parse_FilePassHeader(blob, length - 2, o)
    } else {
        parse_XORObfuscation(blob, length - 2, opts, o)
    }
    return o
}
