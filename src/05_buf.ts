export const has_buf = typeof Buffer !== 'undefined'
    && typeof process !== 'undefined'
    && typeof process.versions !== 'undefined'
    && process.versions.node

export function new_raw_buf(len: number) {
    return new (has_buf ? Buffer : Array)(len)
}

export function s2a(s: string) {
    if (has_buf) {
        return new Buffer(s, 'binary')
    }
    return s.split('').map(function (x) {
        return x.charCodeAt(0) & 0xff
    })
}

export let bconcat = function (bufs) {
    return [].concat(...bufs)
}

export const chr0 = /\u0000/g
export const chr1 = /[\u0001-\u0006]/
