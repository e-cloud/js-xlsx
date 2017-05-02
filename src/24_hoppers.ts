import { __toBuffer, is_buf, new_buf, prep_blob } from './23_binutils'
import { evert_RE, XLSBRecordEnum } from './77_parsetab'

/* [MS-XLSB] 2.1.4 Record */
export function recordhopper(data, cb /*:RecordHopperCB*/, opts? /*:?any*/) {
    if (!data) return
    let tmpbyte
    let cntbyte
    let length
    prep_blob(data, data.l || 0)
    const L = data.length
    let RT = 0
    let tgt = 0
    while (data.l < L) {
        RT = data.read_shift(1)
        if (RT & 0x80) {
            RT = (RT & 0x7F)
                + ((data.read_shift(1) & 0x7F) << 7)
        }
        const R = XLSBRecordEnum[RT] || XLSBRecordEnum[0xFFFF]
        tmpbyte = data.read_shift(1)
        length = tmpbyte & 0x7F
        for (cntbyte = 1; cntbyte < 4 && tmpbyte & 0x80; ++cntbyte) {
            length += ((tmpbyte = data.read_shift(1)) & 0x7F) << 7 * cntbyte
        }
        tgt = data.l + length
        const d = R.f(data, length, opts)
        data.l = tgt
        if (cb(d, R.n, RT)) return
    }
}

/* control buffer usage for fixed-length buffers */
export function buf_array() /*:BufArray*/ {
    const bufs = []
    const blksz = 2048
    const newblk = function ba_newblk(sz) {
        const o /*:Block*/ = new_buf(sz)
        /*:any*/
        prep_blob(o, 0)
        return o
    }

    let curbuf = newblk(blksz)

    const endbuf = function ba_endbuf() {
        if (!curbuf) return
        if (curbuf.length > curbuf.l) curbuf = curbuf.slice(0, curbuf.l)
        if (curbuf.length > 0) bufs.push(curbuf)
        curbuf = null
    }

    const next = function ba_next(sz) {
        if (curbuf && sz < curbuf.length - curbuf.l) return curbuf
        endbuf()
        return curbuf = newblk(Math.max(sz + 1, blksz))
    }

    const end = function ba_end() {
        endbuf()
        return __toBuffer([bufs])
    }

    const push = function ba_push(buf) {
        endbuf()
        curbuf = buf
        next(blksz)
    }

    return {next, push, end, _bufs: bufs}
    /*:any*/
}

export function write_record(ba /*:BufArray*/, type /*:string*/, payload, length? /*:?number*/) {
    const t /*:number*/ = Number(evert_RE[type])
    let l
    if (isNaN(t)) return // TODO: throw something here?
    if (!length) length = XLSBRecordEnum[t].p || (payload || []).length || 0
    l = 1 + (t >= 0x80 ? 1 : 0) + 1 + length
    if (length >= 0x80) ++l
    if (length >= 0x4000) ++l
    if (length >= 0x200000) ++l
    const o = ba.next(l)
    if (t <= 0x7F) {
        o.write_shift(1, t)
    } else {
        o.write_shift(1, (t & 0x7F) + 0x80)
        o.write_shift(1, t >> 7)
    }
    for (let i = 0; i != 4; ++i) {
        if (length >= 0x80) {
            o.write_shift(1, (length & 0x7F) + 0x80)
            length >>= 7
        } else {
            o.write_shift(1, length)
            break
        }
    }
    if (/*:: length != null &&*/length > 0 && is_buf(payload)) ba.push(payload)
}
