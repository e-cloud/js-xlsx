import { keys } from './20_jsutils'
import { getzipdata } from './21_ziputils'
import { decode_cell, encode_range, safe_decode_range } from './27_csfutils'
import { RELS } from './31_rels'
import { parse_cmnt } from './74_xmlbin'

RELS.CMNT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'

export function parse_comments(zip, dirComments, sheets, sheetRels, opts) {
    for (let i = 0; i != dirComments.length; ++i) {
        const canonicalpath = dirComments[i]
        const comments = parse_cmnt(getzipdata(zip, canonicalpath.replace(/^\//, ''), true), canonicalpath, opts)
        if (!comments || !comments.length) {
            continue
        }
        // find the sheets targeted by these comments
        const sheetNames = keys(sheets)
        for (let j = 0; j != sheetNames.length; ++j) {
            const sheetName = sheetNames[j]
            const rels = sheetRels[sheetName]
            if (rels) {
                const rel = rels[canonicalpath]
                if (rel) {
                    insertCommentsIntoSheet(sheetName, sheets[sheetName], comments)
                }
            }
        }
    }
}

export function insertCommentsIntoSheet(sheetName, sheet, comments) {
    const dense = Array.isArray(sheet)
    let cell
    let r
    comments.forEach(function (comment) {
        if (dense) {
            r = decode_cell(comment.ref)
            if (!sheet[r.r]) {
                sheet[r.r] = []
            }
            cell = sheet[r.r][r.c]
        } else {
            cell = sheet[comment.ref]
        }
        if (!cell) {
            cell = {}
            if (dense) {
                sheet[r.r][r.c] = cell
            } else {
                sheet[comment.ref] = cell
            }
            const range = safe_decode_range(sheet['!ref'] || 'BDWGO1000001:A1')
            const thisCell = decode_cell(comment.ref)
            if (range.s.r > thisCell.r) {
                range.s.r = thisCell.r
            }
            if (range.e.r < thisCell.r) {
                range.e.r = thisCell.r
            }
            if (range.s.c > thisCell.c) {
                range.s.c = thisCell.c
            }
            if (range.e.c < thisCell.c) {
                range.e.c = thisCell.c
            }
            const encoded = encode_range(range)
            if (encoded !== sheet['!ref']) {
                sheet['!ref'] = encoded
            }
        }

        if (!cell.c) {
            cell.c = []
        }
        const o = { a: comment.author, t: comment.t, r: comment.r }
        if (comment.h) {
            o.h = comment.h
        }
        cell.c.push(o)
    })
}
