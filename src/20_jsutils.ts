import { SSF } from './10_ssf'

export function isval(x /*:?any*/) /*:boolean*/ {
    return x !== undefined && x !== null
}

export function keys(o /*:any*/) /*:Array<any>*/ {
    return Object.keys(o)
}

export function evert_key(obj /*:any*/, key /*:string*/) /*:EvertType*/ {
    const o = []
    /*:any*/
    const K = keys(obj)
    for (let i = 0; i !== K.length; ++i) {
        o[obj[K[i]][key]] = K[i]
    }
    return o
}

export function evert(obj /*:any*/) /*:EvertType*/ {
    const o = []
    /*:any*/
    const K = keys(obj)
    for (let i = 0; i !== K.length; ++i) {
        o[obj[K[i]]] = K[i]
    }
    return o
}

export function evert_num(obj /*:any*/) /*:EvertNumType*/ {
    const o = []
    /*:any*/
    const K = keys(obj)
    for (let i = 0; i !== K.length; ++i) {
        o[obj[K[i]]] = parseInt(K[i], 10)
    }
    return o
}

export function evert_arr(obj /*:any*/) /*:EvertArrType*/ {
    const o /*:EvertArrType*/ = []
    /*:any*/
    const K = keys(obj)
    for (let i = 0; i !== K.length; ++i) {
        if (o[obj[K[i]]] == null) {
            o[obj[K[i]]] = []
        }
        o[obj[K[i]]].push(K[i])
    }
    return o
}

export function datenum(v /*:Date*/, date1904? /*:?boolean*/) /*:number*/ {
    let epoch = v.getTime()
    if (date1904) {
        epoch += 1462 * 24 * 60 * 60 * 1000
    }
    return (epoch + 2209161600000) / (24 * 60 * 60 * 1000)
}

export function numdate(v /*:number*/) /*:Date*/ {
    const date = SSF.parse_date_code(v)
    const val = new Date()
    if (date == null) {
        throw new Error(`Bad Date Code: ${v}`)
    }
    val.setUTCDate(date.d)
    val.setUTCMonth(date.m - 1)
    val.setUTCFullYear(date.y)
    val.setUTCHours(date.H)
    val.setUTCMinutes(date.M)
    val.setUTCSeconds(date.S)
    return val
}

/* ISO 8601 Duration */
export function parse_isodur(s) {
    let sec = 0
    let mt = 0
    let time = false
    const m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/)
    if (!m) {
        throw new Error(`|${s}| is not an ISO8601 Duration`)
    }
    for (let i = 1; i != m.length; ++i) {
        if (!m[i]) {
            continue
        }
        mt = 1
        if (i > 3) {
            time = true
        }
        switch (m[i].substr(m[i].length - 1)) {
            case 'Y':
                throw new Error(`Unsupported ISO Duration Field: ${m[i].substr(m[i].length - 1)}`)
            case 'D':
                mt *= 24
            /* falls through */
            case 'H':
                mt *= 60
            /* falls through */
            case 'M':
                if (!time) {
                    throw new Error('Unsupported ISO Duration Field: M')
                } else {
                    mt *= 60
                }
            /* falls through */
            case 'S':
                break
        }
        sec += mt * parseInt(m[i], 10)
    }
    return sec
}

const good_pd_date = new Date('2017-02-19T19:06:09.000Z')
const good_pd = good_pd_date.getFullYear() == 2017

export function parseDate(str /*:string|Date*/) /*:Date*/ {
    if (good_pd) {
        return new Date(str)
    }
    if (str instanceof Date) {
        return str
    }
    const n = str.match(/\d+/g) || ['2017', '2', '19', '0', '0', '0']
    return new Date(Date.UTC(+n[0], +n[1] - 1, +n[2], +n[3], +n[4], +n[5]))
}

export function cc2str(arr /*:Array<number>*/) /*:string*/ {
    let o = ''
    for (let i = 0; i != arr.length; ++i) {
        o += String.fromCharCode(arr[i])
    }
    return o
}

export function str2cc(str) {
    const o = []
    for (let i = 0; i != str.length; ++i) {
        o.push(str.charCodeAt(i))
    }
    return o
}

export function dup(o /*:any*/) /*:any*/ {
    if (typeof JSON != 'undefined' && !Array.isArray(o)) {
        return JSON.parse(JSON.stringify(o))
    }
    if (typeof o != 'object' || o == null) {
        return o
    }
    const out = {}
    for (const k in o) {
        if (o.hasOwnProperty(k)) {
            out[k] = dup(o[k])
        }
    }
    return out
}

export function fill(c /*:string*/, l /*:number*/) /*:string*/ {
    let o = ''
    while (o.length < l) {
        o += c
    }
    return o
}
