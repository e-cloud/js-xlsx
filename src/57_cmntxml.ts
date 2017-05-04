import { escapexml, parsexmltag, writetag, writextag, XML_HEADER, XMLNS } from './22_xmlutils'
import { decode_cell } from './27_csfutils'
import { parse_si } from './42_sstxml'

/* 18.7 Comments */
export function parse_comments_xml(data /*:string*/, opts) /*:Array<Comment>*/ {
    /* 18.7.6 CT_Comments */
    if (data.match(/<(?:\w+:)?comments *\/>/)) {
        return []
    }
    const authors = []
    const commentList = []
    const authtag = data.match(/<(?:\w+:)?authors>([^\u2603]*)<\/(?:\w+:)?authors>/)
    if (authtag && authtag[1]) {
        authtag[1].split(/<\/\w*:?author>/).forEach(function (x) {
            if (x === '' || x.trim() === '') {
                return
            }
            const a = x.match(/<(?:\w+:)?author[^>]*>(.*)/)
            if (a) {
                authors.push(a[1])
            }
        })
    }
    const cmnttag = data.match(/<(?:\w+:)?commentList>([^\u2603]*)<\/(?:\w+:)?commentList>/)
    if (cmnttag && cmnttag[1]) {
        cmnttag[1].split(/<\/\w*:?comment>/).forEach(function (x, index) {
            if (x === '' || x.trim() === '') {
                return
            }
            const cm = x.match(/<(?:\w+:)?comment[^>]*>/)
            if (!cm) {
                return
            }
            const y = parsexmltag(cm[0])
            const comment /*:Comment*/ = {
                author: y.authorId && authors[y.authorId] ? authors[y.authorId] : 'sheetjsghost',
                ref: y.ref,
                guid: y.guid,
            }
            /*:any*/
            const cell = decode_cell(y.ref)
            if (opts.sheetRows && opts.sheetRows <= cell.r) {
                return
            }
            const textMatch = x.match(/<(?:\w+:)?text>([^\u2603]*)<\/(?:\w+:)?text>/)
            const rt = !!textMatch && !!textMatch[1] && parse_si(textMatch[1]) || { r: '', t: '', h: '' }
            comment.r = rt.r
            if (rt.r == '<t></t>') {
                rt.t = rt.h = ''
            }
            comment.t = rt.t.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
            if (opts.cellHTML) {
                comment.h = rt.h
            }
            commentList.push(comment)
        })
    }
    return commentList
}

const CMNT_XML_ROOT = writextag('comments', null, { 'xmlns': XMLNS.main[0] })
export function write_comments_xml(data, opts) {
    const o = [XML_HEADER, CMNT_XML_ROOT]

    const iauthor = []
    o.push('<authors>')
    data.map(function (x) {
        return x[1]
    }).forEach(function (comment) {
        comment.map(function (x) {
            return escapexml(x.a)
        }).forEach(function (a) {
            if (iauthor.includes(a)) {
                return
            }
            iauthor.push(a)
            o.push(`<author>${a}</author>`)
        })
    })
    o.push('</authors>')
    o.push('<commentList>')
    data.forEach(function (d) {
        d[1].forEach(function (c) {
            /* 18.7.3 CT_Comment */
            o.push(`<comment ref="${d[0]}" authorId="${iauthor.indexOf(escapexml(c.a))}"><text>`)
            o.push(writetag('t', c.t == null ? '' : c.t))
            o.push('</text></comment>')
        })
    })
    o.push('</commentList>')
    if (o.length > 2) {
        o[o.length] = '</comments>'
        o[1] = o[1].replace('/>', '>')
    }
    return o.join('')
}
