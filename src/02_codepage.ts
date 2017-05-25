import * as cptable from 'codepage/dist/cpexcel.full.js'

export let current_codepage = 1200

function reset_cp() {
    setCurrentCodepage(1200)
}

let setCurrentCodepage = function (cp) {
    current_codepage = cp
}

export function char_codes(data) {
    const o = []
    for (let i = 0, len = data.length; i < len; ++i) {
        o[i] = data.charCodeAt(i)
    }
    return o
}

let debom = function (data: string): string {
    const c1 = data.charCodeAt(0)
    const c2 = data.charCodeAt(1)
    if (c1 == 0xFF && c2 == 0xFE) {
        return data.substr(2)
    }
    if (c1 == 0xFE && c2 == 0xFF) {
        return data.substr(2)
    }
    if (c1 == 0xFEFF) {
        return data.substr(1)
    }
    return data
}

let _getchar = function _gc1(x) {
    return String.fromCharCode(x)
}

if (typeof cptable !== 'undefined') {
    setCurrentCodepage = function (cp) {
        current_codepage = cp
    }
    debom = function (data) {
        if (data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) {
            return cptable.utils.decode(1200, char_codes(data.substr(2)))
        }
        return data
    }
    _getchar = function _gc2(x) {
        if (current_codepage === 1200) {
            return String.fromCharCode(x)
        }
        return cptable.utils.decode(current_codepage, [x & 255, x >> 8])[0]
    }
}

export {
    reset_cp,
    setCurrentCodepage,
    debom,
    _getchar,
}
