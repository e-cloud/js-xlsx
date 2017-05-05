import XLSX = require('../');
import testCommon = require('./Common.js');

const file = '\u05D7\u05D9\u05E9\u05D5\u05D1_\u05E0\u05E7\u05D5\u05D3\u05D5\u05EA_\u05D6\u05D9\u05DB\u05D5\u05D9.xlsx'

describe(file, function () {
    testCommon(file)
})
