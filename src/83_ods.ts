/* Part 3: Packages */
import * as JSZip from 'jszip'
import { getzipdata, getzipstr, safegetzipfile } from './21_ziputils'
import { utf8read } from './22_xmlutils'
import { parse_manifest, write_manifest, write_rdf } from './32_odmanrdf'
import { parse_content_xml } from './80_parseods'
import { write_content_xml } from './81_writeods'

export function parse_ods(zip: ZIPFile, opts ?: ParseOpts = {}) {
    const ods = !!safegetzipfile(zip, 'objectdata')
    if (ods) {
        // todo: remove var
        var manifest = parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts)
    }
    const content = getzipstr(zip, 'content.xml')
    if (!content) {
        throw new Error(`Missing content.xml in ${ods ? 'ODS' : 'UOF'} file`)
    }
    return parse_content_xml(ods ? content : utf8read(content), opts)
}

export function parse_fods(data: string, opts ?: ParseOpts) {
    return parse_content_xml(data, opts)
}

export function write_ods(wb, opts) {
    if (opts.bookType == 'fods') {
        return write_content_xml(wb, opts)
    }

    /*:: if(!jszip) throw new Error("JSZip is not available"); */
    const zip = new JSZip()
    let f = ''

    const manifest: Array<Array<string>> = []
    const rdf = []

    /* 3:3.3 and 2:2.2.4 */
    f = 'mimetype'
    zip.file(f, 'application/vnd.oasis.opendocument.spreadsheet')

    /* Part 1 Section 2.2 Documents */
    f = 'content.xml'
    zip.file(f, write_content_xml(wb, opts))
    manifest.push([f, 'text/xml'])
    rdf.push([f, 'ContentFile'])

    /* Part 3 Section 6 Metadata Manifest File */
    f = 'manifest.rdf'
    zip.file(f, write_rdf(rdf, opts))
    manifest.push([f, 'application/rdf+xml'])

    /* Part 3 Section 4 Manifest File */
    f = 'META-INF/manifest.xml'
    zip.file(f, write_manifest(manifest, opts))

    return zip
}
