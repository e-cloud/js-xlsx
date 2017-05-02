/* 18.2.28 (CT_WorkbookProtection) Defaults */
import { parsexmlbool } from './22_xmlutils'
import { _ssfopts } from './66_wscommon'
export const WBPropsDef = [
    ['allowRefreshQuery', '0'],
    ['autoCompressPictures', '1'],
    ['backupFile', '0'],
    ['checkCompatibility', '0'],
    ['codeName', ''],
    ['date1904', '0'],
    ['dateCompatibility', '1'],
    //['defaultThemeVersion', '0'],
    ['filterPrivacy', '0'],
    ['hidePivotFieldList', '0'],
    ['promptedSolutions', '0'],
    ['publishItems', '0'],
    ['refreshAllConnections', false],
    ['saveExternalLinkValues', '1'],
    ['showBorderUnselectedTables', '1'],
    ['showInkAnnotation', '1'],
    ['showObjects', 'all'],
    ['showPivotChartFilter', '0'],
    //['updateLinks', 'userSet']
]

/* 18.2.30 (CT_BookView) Defaults */
export const WBViewDef = [
    ['activeTab', '0'],
    ['autoFilterDateGrouping', '1'],
    ['firstSheet', '0'],
    ['minimized', '0'],
    ['showHorizontalScroll', '1'],
    ['showSheetTabs', '1'],
    ['showVerticalScroll', '1'],
    ['tabRatio', '600'],
    ['visibility', 'visible'],
    //window{Height,Width}, {x,y}Window
]

/* 18.2.19 (CT_Sheet) Defaults */
export const SheetDef = [
    //['state', 'visible']
]

/* 18.2.2  (CT_CalcPr) Defaults */
export const CalcPrDef = [
    ['calcCompleted', 'true'],
    ['calcMode', 'auto'],
    ['calcOnSave', 'true'],
    ['concurrentCalc', 'true'],
    ['fullCalcOnLoad', 'false'],
    ['fullPrecision', 'true'],
    ['iterate', 'false'],
    ['iterateCount', '100'],
    ['iterateDelta', '0.001'],
    ['refMode', 'A1'],
]

/* 18.2.3 (CT_CustomWorkbookView) Defaults */
export const CustomWBViewDef = [
    ['autoUpdate', 'false'],
    ['changesSavedWin', 'false'],
    ['includeHiddenRowCol', 'true'],
    ['includePrintSettings', 'true'],
    ['maximized', 'false'],
    ['minimized', 'false'],
    ['onlySync', 'false'],
    ['personalView', 'false'],
    ['showComments', 'commIndicator'],
    ['showFormulaBar', 'true'],
    ['showHorizontalScroll', 'true'],
    ['showObjects', 'all'],
    ['showSheetTabs', 'true'],
    ['showStatusbar', 'true'],
    ['showVerticalScroll', 'true'],
    ['tabRatio', '600'],
    ['xWindow', '0'],
    ['yWindow', '0'],
]

export function push_defaults_array(target, defaults) {
    for (let j = 0; j != target.length; ++j) {
        const w = target[j]
        for (let i = 0; i != defaults.length; ++i) {
            const z = defaults[i]
            if (w[z[0]] == null) {
                w[z[0]] = z[1]
            }
        }
    }
}
export function push_defaults(target, defaults) {
    for (let i = 0; i != defaults.length; ++i) {
        const z = defaults[i]
        if (target[z[0]] == null) {
            target[z[0]] = z[1]
        }
    }
}

export function parse_wb_defaults(wb) {
    push_defaults(wb.WBProps, WBPropsDef)
    push_defaults(wb.CalcPr, CalcPrDef)

    push_defaults_array(wb.WBView, WBViewDef)
    push_defaults_array(wb.Sheets, SheetDef)

    _ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904, 'date1904')
}

export function check_wb_names(N) {
    const badchars = '][*?/\\'.split('')
    N.forEach(function (n, i) {
        badchars.forEach(function (c) {
            if (n.includes(c)) {
                throw new Error('Sheet name cannot contain : \\ / ? * [ ]')
            }
        })
        if (n.length > 31) throw new Error('Sheet names cannot exceed 31 chars')
        for (let j = 0; j < i; ++j) {
            if (n == N[j]) {
                throw new Error(`Duplicate Sheet Name: ${n}`)
            }
        }
    })
}
export function check_wb(wb) {
    if (!wb || !wb.SheetNames || !wb.Sheets) {
        throw new Error('Invalid Workbook')
    }
    check_wb_names(wb.SheetNames)
    /* TODO: validate workbook */
}
