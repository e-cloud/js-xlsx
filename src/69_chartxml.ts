import { unescapexml } from './22_xmlutils'
import { encode_col, encode_range, encode_row } from './27_csfutils'

function parse_numCache(data) {
    const col = [];

    /* 21.2.2.150 pt CT_NumVal */
    (data.match(/<c:pt idx="(\d*)">(.*?)<\/c:pt>/mg) || []).forEach(function (pt) {
        const q = pt.match(/<c:pt idx="(.*?)"><c:v>(.*)<\/c:v><\/c:pt>/)
        if (!q) return
        col[+q[1]] = +q[2]
    })

    /* 21.2.2.71 formatCode CT_Xstring */
    const nf = unescapexml((data.match(/<c:formatCode>(.*?)<\/c:formatCode>/) || ['', 'General'])[1])

    return [col, nf]
}

/* 21.2 DrawingML - Charts */
export function parse_chart(data, name /*:string*/, opts, rels, wb, csheet) {
    const cs = csheet || {'!type': 'chart'}
    if (!data) return csheet

    /* 21.2.2.27 chart CT_Chart */

    let C = 0

    let R = 0
    let col = 'A'
    const refguess = {s: {r: 2000000, c: 2000000}, e: {r: 0, c: 0}};

    /* 21.2.2.120 numCache CT_NumData */
    (data.match(/<c:numCache>.*?<\/c:numCache>/gm) || []).forEach(function (nc) {
        const cache = parse_numCache(nc)
        refguess.s.r = refguess.s.c = 0
        refguess.e.c = C
        col = encode_col(C)
        cache[0].forEach(function (n, i) {
            cs[col + encode_row(i)] = {t: 'n', v: n, z: cache[1]}
            R = i
        })
        if (refguess.e.r < R) refguess.e.r = R
        ++C
    })
    if (C > 0) cs['!ref'] = encode_range(refguess)
    return cs
}
