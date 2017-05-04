function hex2RGB(h) {
    const o = h.substr(h[0] === '#' ? 1 : 0, 6)
    return [parseInt(o.substr(0, 2), 16), parseInt(o.substr(2, 2), 16), parseInt(o.substr(4, 2), 16)]
}

export function rgb2Hex(rgb) {
    let o = 1
    for (let i = 0; i != 3; ++i) {
        o = o * 256 + (rgb[i] > 255 ? 255 : rgb[i] < 0 ? 0 : rgb[i])
    }
    return o.toString(16).toUpperCase().substr(1)
}

function rgb2HSL(rgb) {
    const R = rgb[0] / 255
    const G = rgb[1] / 255
    const B = rgb[2] / 255
    const M = Math.max(R, G, B)
    const m = Math.min(R, G, B)
    const C = M - m
    if (C === 0) {
        return [0, 0, R]
    }

    let H6 = 0
    let S = 0
    const L2 = M + m
    S = C / (L2 > 1 ? 2 - L2 : L2)
    switch (M) {
        case R:
            H6 = ((G - B) / C + 6) % 6
            break
        case G:
            H6 = (B - R) / C + 2
            break
        case B:
            H6 = (R - G) / C + 4
            break
    }
    return [H6 / 6, S, L2 / 2]
}

function hsl2RGB(hsl) {
    const H = hsl[0]
    const S = hsl[1]
    const L = hsl[2]
    const C = S * 2 * (L < 0.5 ? L : 1 - L)
    const m = L - C / 2
    const rgb = [m, m, m]
    const h6 = 6 * H

    let X
    if (S !== 0) {
        switch (h6 | 0) {
            case 0:
            case 6:
                X = C * h6
                rgb[0] += C
                rgb[1] += X
                break
            case 1:
                X = C * (2 - h6)
                rgb[0] += X
                rgb[1] += C
                break
            case 2:
                X = C * (h6 - 2)
                rgb[1] += C
                rgb[2] += X
                break
            case 3:
                X = C * (4 - h6)
                rgb[1] += X
                rgb[2] += C
                break
            case 4:
                X = C * (h6 - 4)
                rgb[2] += C
                rgb[0] += X
                break
            case 5:
                X = C * (6 - h6)
                rgb[2] += X
                rgb[0] += C
                break
        }
    }
    for (let i = 0; i != 3; ++i) {
        rgb[i] = Math.round(rgb[i] * 255)
    }
    return rgb
}

/* 18.8.3 bgColor tint algorithm */
export function rgb_tint(hex, tint) {
    if (tint === 0) {
        return hex
    }
    const hsl = rgb2HSL(hex2RGB(hex))
    if (tint < 0) {
        hsl[2] = hsl[2] * (1 + tint)
    } else {
        hsl[2] = 1 - (1 - hsl[2]) * (1 - tint)
    }
    return rgb2Hex(hsl2RGB(hsl))
}

/* 18.3.1.13 width calculations */
/* [MS-OI29500] 2.1.595 Column Width & Formatting */
export const DEF_MDW = 6

const MAX_MDW = 15
const MIN_MDW = 1
export let MDW = DEF_MDW

export function setMDW(val) {
    MDW = val
}

export function width2px(width) {
    return Math.floor((width + Math.round(128 / MDW) / 256) * MDW)
}

export function px2char(px) {
    return Math.floor((px - 5) / MDW * 100 + 0.5) / 100
}

export function char2width(chr) {
    return Math.round((chr * MDW + 5) / MDW * 256) / 256
}

function px2char_(px) {
    return ((px - 5) / MDW * 100 + 0.5) / 100
}

function char2width_(chr) {
    return (chr * MDW + 5) / MDW * 256 / 256
}

function cycle_width(collw) {
    return char2width(px2char(width2px(collw)))
}

/* XLSX/XLSB/XLS specify width in units of MDW */
export function find_mdw_colw(collw) {
    let delta = Infinity
    let _MDW = MIN_MDW
    for (MDW = MIN_MDW; MDW < MAX_MDW; ++MDW) {
        if (Math.abs(collw - cycle_width(collw)) <= delta) {
            delta = Math.abs(collw - cycle_width(collw))
            _MDW = MDW
        }
    }
    MDW = _MDW
}

/* XLML specifies width in terms of pixels */
export function find_mdw_wpx(wpx) {
    let delta = Infinity
    let guess = 0
    let _MDW = MIN_MDW
    for (MDW = MIN_MDW; MDW < MAX_MDW; ++MDW) {
        guess = char2width_(px2char_(wpx)) * 256
        guess = guess % 1
        if (guess > 0.5) {
            guess--
        }
        if (Math.abs(guess) < delta) {
            delta = Math.abs(guess)
            _MDW = MDW
        }
    }
    MDW = _MDW
}

export function process_col(coll /*:ColInfo*/) {
    if (coll.width) {
        coll.wpx = width2px(coll.width)
        coll.wch = px2char(coll.wpx)
        coll.MDW = MDW
    } else if (coll.wpx) {
        coll.wch = px2char(coll.wpx)
        coll.width = char2width(coll.wch)
        coll.MDW = MDW
    } else if (typeof coll.wch == 'number') {
        coll.width = char2width(coll.wch)
        coll.wpx = width2px(coll.width)
        coll.MDW = MDW
    }
    if (coll.customWidth) {
        delete coll.customWidth
    }
}

const DEF_PPI = 96
const PPI = DEF_PPI

export function px2pt(px) {
    return px * 96 / PPI
}

export function pt2px(pt) {
    return pt * PPI / 96
}

/* [MS-EXSPXML3] 2.4.54 ST_enmPattern */
export const XLMLPatternTypeMap = {
    'None': 'none',
    'Solid': 'solid',
    'Gray50': 'mediumGray',
    'Gray75': 'darkGray',
    'Gray25': 'lightGray',
    'HorzStripe': 'darkHorizontal',
    'VertStripe': 'darkVertical',
    'ReverseDiagStripe': 'darkDown',
    'DiagStripe': 'darkUp',
    'DiagCross': 'darkGrid',
    'ThickDiagCross': 'darkTrellis',
    'ThinHorzStripe': 'lightHorizontal',
    'ThinVertStripe': 'lightVertical',
    'ThinReverseDiagStripe': 'lightDown',
    'ThinHorzCross': 'lightGrid',
}
