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

// https://fable.io/docs/javascript/features.html#paramobject
[<AllowNullLiteral>]
[<Global>]
type TableColumn
    [<ParamObject; Emit("$0")>]
    (   
        name: string, 
        ?filterButton: bool, 
        ?totalsRowLabel: TotalsFunctions, 
        ?totalsRowFunction: string, 
        ?totalsRowFormula: obj
    ) =
    member val name = jsNative with get, set
    member val filterButton = jsNative with get, set
    member val totalsRowLabel = jsNative with get, set
    member val totalsRowFunction = jsNative with get, set
    member val totalsRowFormula = jsNative with get, set

[<AllowNullLiteral>]
[<Global>]
type Table
    [<ParamObject; Emit("$0")>]
    (   name: string, 
        ref: CellAdress, 
        columns: TableColumn [], 
        rows: RowValues [] [], 
        ?displayName: string, 
        ?headerRow: bool, 
        ?totalsRow: bool, 
        ?style: obj
    ) =
    //interface ITable with
    member val name = jsNative with get, set
    member val displayName = jsNative with get, set
    member val ref = jsNative with get, set
    member val headerRow = jsNative with get, set
    member val totalsRow = jsNative with get, set
    member val style = jsNative with get, set
    member val columns = jsNative with get, set
    member val rows = jsNative with get, set
    member val tableRef = jsNative with get