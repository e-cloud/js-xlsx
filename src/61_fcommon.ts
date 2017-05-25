/* TODO: it will be useful to parse the function str */
import { decode_cell, decode_col, decode_range, decode_row, encode_col, encode_row } from './27_csfutils'

export const rc_to_a1 = function () {
    const rcregex = /(^|[^A-Za-z])R(\[?)(-?\d+|)\]?C(\[?)(-?\d+|)\]?/g
    let rcbase: Cell = { r: 0, c: 0 }

    function rcfunc($$, $1, $2, $3, $4, $5) {
        let R = $3.length > 0 ? parseInt($3, 10) | 0 : 0
        let C = $5.length > 0 ? parseInt($5, 10) | 0 : 0
        if (C < 0 && $4.length === 0) {
            C = 0
        }
        let cRel = false
        let rRel = false
        if ($4.length > 0 || $5.length == 0) {
            cRel = true
        }
        if (cRel) {
            C += rcbase.c
        } else {
            --C
        }
        if ($2.length > 0 || $3.length == 0) {
            rRel = true
        }
        if (rRel) {
            R += rcbase.r
        } else {
            --R
        }
        return $1 + (cRel ? '' : '$') + encode_col(C) + (rRel ? '' : '$') + encode_row(R)
    }

    return function rc_to_a1(fstr: string, base: Cell): string {
        rcbase = base
        return fstr.replace(rcregex, rcfunc)
    }
}()

export const crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g
export const a1_to_rc = function () {
    return function a1_to_rc(fstr, base) {
        return fstr.replace(crefregex, function ($0, $1, $2, $3, $4, $5, off, str) {
            /* TODO: handle fixcol / fixrow */
            const c = decode_col($3) - base.c
            const r = decode_row($5) - base.r
            return `${$1}R${r == 0 ? '' : `[${r}]`}C${c == 0 ? '' : `[${c}]`}`
        })
    }
}()

/* no defined name can collide with a valid cell address A1:XFD1048576 ... except LOG10! */
export function shift_formula_str(f: string, delta: Cell): string {
    return f.replace(crefregex, function ($0, $1, $2, $3, $4, $5, off, str) {
        return $1 + ($2 == '$' ? $2 + $3 : encode_col(decode_col($3) + delta.c)) + ($4 == '$'
            ? $4 + $5
            : encode_row(decode_row($5) + delta.r))
    })
}

export function shift_formula_xlsx(f: string, range: string, cell: string): string {
    const r = decode_range(range)
    const s = r.s
    const c = decode_cell(cell)
    const delta = { r: c.r - s.r, c: c.c - s.c }
    return shift_formula_str(f, delta)
}
