import XLSX = require('../');
import testCommon = require('./Common.js');

const file = 'mixed_sheets.xlsx'

describe(file, function () {
    testCommon(file)
})
