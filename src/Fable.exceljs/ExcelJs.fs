module Fable.ExcelJs.ExcelJs

open Fable.Core
open Fable.Core.JsInterop

type ExcelJS =
    [<Emit("new $0.Workbook()")>]
    abstract member Workbook: unit -> Workbook
    
let exceljs: ExcelJS = importDefault "exceljs"
