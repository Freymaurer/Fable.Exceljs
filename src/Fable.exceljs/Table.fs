[<AutoOpen>]
module Fable.ExcelJs.Table

open Fable.Core
open Fable.Core.JsInterop
open Fable.ExcelJs
open Cell

[<RequireQualifiedAccess>]
type TotalsFunctions =
/// Default
| [<CompiledName("Total")>] Total
/// No totals function for this column
| [<CompiledName("none")>] None
/// Compute average for the column
| [<CompiledName("average")>] Average
/// Count the entries that are numbers
| [<CompiledName("countNums")>] CountNums
/// Count of entries
| [<CompiledName("count")>] Count
/// The maximum value in this column
| [<CompiledName("max")>] Max
/// The minimum value in this column
| [<CompiledName("min")>] Min
/// The standard deviation for this column
| [<CompiledName("stdDev")>] StdDev
/// The variance for this column
| [<CompiledName("var")>] Var
/// The sum of entries for this column
| [<CompiledName("sum")>] Sum
/// A custom formula. Requires an associated totalsRowFormula value.
| [<CompiledName("custom")>] Custom

with
    static member defaultValue = Total

type ITableColumn = 
    /// The name of the column, also used in the header
    abstract member name: string with get, set
    /// Switches the filter control in the header. Default Value = false
    abstract member filterButton: bool with get, set
    /// Label to describe the totals row (first column). Default Value = Total
    abstract member totalsRowLabel: TotalsFunctions with get, set
    /// Name of the totals function. Default Value = "None"
    abstract member totalsRowFunction: string with get, set
    /// Optional formula for custom functions
    abstract member totalsRowFormula: obj with get, set

type TableColumn(name: string, ?filterButton: bool, ?totalsRowLabel: TotalsFunctions, ?totalsRowFunction: string, ?totalsRowFormula: obj) =
    interface ITableColumn with
        member val name = name with get, set
        member val filterButton = Option.defaultValue false filterButton with get, set
        member val totalsRowLabel = Option.defaultValue TotalsFunctions.defaultValue totalsRowLabel with get, set
        member val totalsRowFunction = Option.defaultValue "none" totalsRowFunction with get, set
        member val totalsRowFormula = Option.defaultValue (box None) totalsRowFormula with get, set

//https://github.com/exceljs/exceljs#modifying-tables

// https://github.com/exceljs/exceljs#table-properties
type ITable =
    /// The name of the table
    abstract member name: string with get, set
    /// The display name of the table. Default Value = "name"
    abstract member displayName: string with get, set
    /// Top left cell of the table
    abstract member ref: CellAdress with get, set
    /// Show headers at top of table
    abstract member headerRow: bool with get, set
    /// Show totals at bottom of table
    abstract member totalsRow: bool with get, set
    /// Extra style properties
    abstract member style: obj with get, set
    /// Column definitions
    abstract member columns: TableColumn [] with get, set
    /// Rows of data
    abstract member rows: RowValues [] [] with get, set
    //abstract member tableRef: CellRange with get // how to do this?

type Table(name: string, ref: CellAdress, columns: TableColumn [], rows: RowValues [] [], ?displayName: string, ?headerRow: bool, ?totalsRow: bool, ?style: obj) =
    interface ITable with
        member val name = name with get, set
        member val displayName = Option.defaultValue "name" displayName with get, set
        member val ref = ref with get, set
        member val headerRow = Option.defaultValue true headerRow with get, set
        member val totalsRow = Option.defaultValue false totalsRow with get, set
        member val style = Option.defaultValue (box None) style with get, set
        member val columns = columns with get, set
        member val rows = rows with get, set