import { char_codes, debom } from './02_codepage'
import { has_buf } from './05_buf'
import { cc2str, keys, str2cc } from './20_jsutils'

export function getdatastr(data) /*:?string*/ {
    if (!data) {
        return null
    }
    if (data.data) {
        return debom(data.data)
    }
    if (data.asNodeBuffer && has_buf) {
        return debom(data.asNodeBuffer().toString('binary'))
    }
    if (data.asBinary) {
        return debom(data.asBinary())
    }
    if (data._data && data._data.getContent) {
        return debom(cc2str(Array.prototype.slice.call(data._data.getContent(), 0)))
    }
    return null
}

export function getdatabin(data) {
    if (!data) {
        return null
    }
    if (data.data) {
        return char_codes(data.data)
    }
    if (data.asNodeBuffer && has_buf) {
        return data.asNodeBuffer()
    }
    if (data._data && data._data.getContent) {
        const o = data._data.getContent()
        if (typeof o == 'string') {
            return str2cc(o)
        }
        return Array.prototype.slice.call(o)
    }
    return null
}

export function getdata(data) {
    return data && data.name.slice(-4) === '.bin' ? getdatabin(data) : getdatastr(data)
}

/* Part 2 Section 10.1.2 "Mapping Content Types" Names are case-insensitive */
/* OASIS does not comment on filename case sensitivity */
export function safegetzipfile(zip, file /*:string*/) {
    const k = keys(zip.files)
    const f = file.toLowerCase()
    const g = f.replace(/\//g, '\\')
    for (let i = 0; i < k.length; ++i) {
        const n = k[i].toLowerCase()
        if (f == n || g == n) {
            return zip.files[k[i]]
        }
    }
    return null
}

export function getzipfile(zip, file /*:string*/) {
    const o = safegetzipfile(zip, file)
    if (o == null) {
        throw new Error(`Cannot find file ${file} in zip`)
    }
    return o
}

export function getzipdata(zip, file /*:string*/, safe? /*:?boolean*/) {
    if (!safe) {
        return getdata(getzipfile(zip, file))
    }
    if (!file) {
        return null
    }
    try {
        return getzipdata(zip, file)
    } catch (e) {
        return null
    }
}

export function getzipstr(zip, file /*:string*/, safe? /*:?boolean*/) /*:?string*/ {
    if (!safe) {
        return getdatastr(getzipfile(zip, file))
    }
    if (!file) {
        return null
    }
    try {
        return getzipstr(zip, file)
    } catch (e) {
        return null
    }
}

export function resolve_path(path /*:string*/, base /*:string*/) /*:string*/ {
    const result = base.split('/')
    if (base.slice(-1) != '/') {
        result.pop()
    } // folder path
    const target = path.split('/')
    while (target.length !== 0) {
        const step = target.shift()
        if (step === '..') {
            result.pop()
        } else if (step !== '.') {
            result.push(step)
        }
    }
    return result.join('/')
}
