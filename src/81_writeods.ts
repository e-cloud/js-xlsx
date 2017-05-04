import { parseDate } from './20_jsutils'
import { escapexml, writextag, wxt_helper, XML_HEADER } from './22_xmlutils'
import { decode_range, encode_cell } from './27_csfutils'
import { csf_to_ods_formula } from './65_fods'

export const write_content_xml /*:{(wb:any, opts:any):string}*/ = function () {
    const null_cell_xml = '          <table:table-cell />\n'
    const covered_cell_xml = '          <table:covered-table-cell/>\n'
    const write_ws = function (ws, wb, i /*:number*/, opts) /*:string*/ {
        /* Section 9 Tables */
        const o = []
        o.push(`      <table:table table:name="${escapexml(wb.SheetNames[i])}">\n`)
        let R = 0
        let C = 0
        const range = decode_range(ws['!ref'])
        const marr = ws['!merges'] || []
        let mi = 0
        const dense = Array.isArray(ws)
        for (R = 0; R < range.s.r; ++R) {
            o.push('        <table:table-row></table:table-row>\n')
        }
        for (; R <= range.e.r; ++R) {
            o.push('        <table:table-row>\n')
            for (C = 0; C < range.s.c; ++C) {
                o.push(null_cell_xml)
            }
            for (; C <= range.e.c; ++C) {
                let skip = false
                const ct = {}
                let textp = ''
                for (mi = 0; mi != marr.length; ++mi) {
                    if (marr[mi].s.c > C) {
                        continue
                    }
                    if (marr[mi].s.r > R) {
                        continue
                    }
                    if (marr[mi].e.c < C) {
                        continue
                    }
                    if (marr[mi].e.r < R) {
                        continue
                    }
                    if (marr[mi].s.c != C || marr[mi].s.r != R) {
                        skip = true
                    }
                    ct['table:number-columns-spanned'] = (marr[mi].e.c - marr[mi].s.c + 1)
                    ct['table:number-rows-spanned'] = (marr[mi].e.r - marr[mi].s.r + 1)
                    break
                }
                if (skip) {
                    o.push(covered_cell_xml)
                    continue
                }
                const ref = encode_cell({ r: R, c: C })
                const cell = dense ? (ws[R] || [])[C] : ws[ref]
                if (cell && cell.f) {
                    ct['table:formula'] = escapexml(csf_to_ods_formula(cell.f))
                    if (cell.F) {
                        if (cell.F.substr(0, ref.length) == ref) {
                            const _Fref = decode_range(cell.F)
                            ct['table:number-matrix-columns-spanned'] = (_Fref.e.c - _Fref.s.c + 1)
                            ct['table:number-matrix-rows-spanned'] = (_Fref.e.r - _Fref.s.r + 1)
                        }
                    }
                }
                if (!cell) {
                    o.push(null_cell_xml)
                    continue
                }
                switch (cell.t) {
                    case 'b':
                        textp = (cell.v ? 'TRUE' : 'FALSE')
                        ct['office:value-type'] = 'boolean'
                        ct['office:boolean-value'] = (cell.v ? 'true' : 'false')
                        break
                    case 'n':
                        textp = (cell.w || String(cell.v || 0))
                        ct['office:value-type'] = 'float'
                        ct['office:value'] = (cell.v || 0)
                        break
                    case 's':
                    case 'str':
                        textp = escapexml(cell.v)
                        ct['office:value-type'] = 'string'
                        break
                    case 'd':
                        textp = (cell.w || (parseDate(cell.v).toISOString()))
                        ct['office:value-type'] = 'date'
                        ct['office:date-value'] = (parseDate(cell.v).toISOString())
                        ct['table:style-name'] = 'ce1'
                        break
                    //case 'e':
                    default:
                        o.push(null_cell_xml)
                        continue
                }
                o.push(writextag('table:table-cell', writextag('text:p', textp, {}), ct))
            }
            o.push('        </table:table-row>\n')
        }
        o.push('      </table:table>\n')
        return o.join('')
    }

    function write_automatic_styles_ods(o/*:Array<string>*/) {
        o.push(' <office:automatic-styles>\n')
        o.push('  <number:date-style style:name="N37" number:automatic-order="true">\n')
        o.push('   <number:month number:style="long"/>\n')
        o.push('   <number:text>/</number:text>\n')
        o.push('   <number:day number:style="long"/>\n')
        o.push('   <number:text>/</number:text>\n')
        o.push('   <number:year/>\n')
        o.push('  </number:date-style>\n')
        o.push('  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>\n')
        o.push(' </office:automatic-styles>\n')
    }

    return function wcx(wb, opts) {
        const o = [XML_HEADER]
        /* 3.1.3.2 */
        const attr = wxt_helper({
            'xmlns:office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'xmlns:table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
            'xmlns:style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'xmlns:text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'xmlns:draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
            'xmlns:fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
            'xmlns:xlink': 'http://www.w3.org/1999/xlink',
            'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
            'xmlns:meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
            'xmlns:number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
            'xmlns:presentation': 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
            'xmlns:svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
            'xmlns:chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
            'xmlns:dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
            'xmlns:math': 'http://www.w3.org/1998/Math/MathML',
            'xmlns:form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
            'xmlns:script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
            'xmlns:ooo': 'http://openoffice.org/2004/office',
            'xmlns:ooow': 'http://openoffice.org/2004/writer',
            'xmlns:oooc': 'http://openoffice.org/2004/calc',
            'xmlns:dom': 'http://www.w3.org/2001/xml-events',
            'xmlns:xforms': 'http://www.w3.org/2002/xforms',
            'xmlns:xsd': 'http://www.w3.org/2001/XMLSchema',
            'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
            'xmlns:sheet': 'urn:oasis:names:tc:opendocument:sh33tjs:1.0',
            'xmlns:rpt': 'http://openoffice.org/2005/report',
            'xmlns:of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
            'xmlns:xhtml': 'http://www.w3.org/1999/xhtml',
            'xmlns:grddl': 'http://www.w3.org/2003/g/data-view#',
            'xmlns:tableooo': 'http://openoffice.org/2009/table',
            'xmlns:drawooo': 'http://openoffice.org/2010/draw',
            'xmlns:calcext': 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
            'xmlns:loext': 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
            'xmlns:field': 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0',
            'xmlns:formx': 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0',
            'xmlns:css3t': 'http://www.w3.org/TR/css3-text/',
            'office:version': '1.2',
        })

        const fods = wxt_helper({
            'xmlns:config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
            'office:mimetype': 'application/vnd.oasis.opendocument.spreadsheet',
        })

        if (opts.bookType == 'fods') {
            o.push(`<office:document${attr}${fods}>\n`)
        } else {
            o.push(`<office:document-content${attr}>`)
        }
        write_automatic_styles_ods(o)
        o.push('  <office:body>\n')
        o.push('    <office:spreadsheet>\n')
        for (let i = 0; i != wb.SheetNames.length; ++i) {
            o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts))
        }
        o.push('    </office:spreadsheet>\n')
        o.push('  </office:body>\n')
        if (opts.bookType == 'fods') {
            o.push('</office:document>')
        } else {
            o.push('</office:document-content>')
        }
        return o.join('')
    }
}()
