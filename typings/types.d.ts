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
