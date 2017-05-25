declare interface Block {
}

declare interface BufArray {
    newblk(sz: number): Block;

    next(sz: number): Block;

    end(): any;

    push(buf: Block): void;
}

declare type RecordHopperCB = (d: any, Rn: string, RT: number) => boolean;

declare interface EvertType {
    [type: string]: string
}

declare interface EvertNumType {
    [type: string]: number
}

declare interface EvertArrType {
    [type: string]: Array<string>
}

declare interface StringConv {
    (string): string
}

declare interface WriteObjStrFactory {
    from_sheet(ws: Worksheet, o: any): string
}

declare interface ZIPFile {
}

declare interface XLString {
    t: string;
    r?: string;
    h?: string;
}

declare interface WorkbookFile {
}

declare interface Workbook {
    SheetNames: Array<string>;
    Sheets: any;

    Props?: any;
    Custprops?: any;
    Themes?: any;

    Workbook?: WBWBProps;

    SSF?: { [n: number]: string };
    cfb?: any;
}

declare interface WorkbookProps {
    SheetNames?: string[];
}

declare interface WBWBProps {
    Sheets: WBWSProp[];
}

declare interface WBWSProp {
    Hidden?: number;
    name?: string;
}

declare interface CellAddress {
    r: number;
    c: number;
}

declare type CellAddrSpec = CellAddress | string;

declare interface Cell {
    l?: Hyperlink
    t: string
    v: any
    z?: string
}

declare interface Range {
    s: CellAddress;
    e: CellAddress;
}

declare interface Worksheet {
}

declare interface Sheet2CSVOpts {
}

declare interface Sheet2JSONOpts {
}

declare interface ParseOpts {
}

declare interface WriteOpts {
}

declare interface WriteFileOpts {
}

declare interface RawData {
}

declare interface TypeOpts {
    type: string;
}

declare interface XLSXModule {
}

declare interface SST {
    [n: number]: XLString;

    Count: number;
    Unique: number;

    push(x: XLString): void;

    length: number;
}

declare interface Comment {
}

declare interface ColInfo {
    MDW?: number;  // Excel's "Max Digit Width" unit, always integral
    width: number; // width in Excel's "Max Digit Width", width*256 is integral
    wpx?: number;  // width in screen pixels
    wch?: number;  // intermediate character calculation
}

declare type AOA = Array<Array<any>>;

declare type RowInfo = {
    /* visibility */
    hidden?: boolean; // if true, the row is hidden

    /* row height is specified in one of the following ways: */
    hpx?: number;     // height in screen pixels
    hpt?: number;     // height in points
}

declare interface SSFTable {

}

declare interface Margins {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
}

declare interface DefinedName {
    Name: string;
    Ref: string;
    Sheet?: number;
    Comment?: string;
}

declare interface Hyperlink {
    Target: string;
    Tooltip?: string;
}

declare interface Sheet2HTMLOpts {
    editable?: boolean
    dense?: boolean
    header?: boolean
    footer?: boolean
}
