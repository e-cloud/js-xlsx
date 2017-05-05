#!/usr/bin/env node

import fs = require('fs');
const paths = fs.readFileSync('tests/fixtures.lst', 'utf-8').split('\n')
const aux = ['multiformat.lst', './misc/ssf.json', './test_files/biff5/number_format_greek.xls']
const fullpaths = paths.concat(aux)

fs.writeFileSync('tests/fixtures.js', fullpaths.map(function (x) {
        return [x, fs.existsSync(x) ? fs.readFileSync(x).toString('base64') : '']
    })
    .map(function (w) {
        return `fs['${w[0]}'] = '${w[1]}';\n`
    })
    .join('')
)
