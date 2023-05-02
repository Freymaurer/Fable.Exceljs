[<AutoOpen>]
module Fable.ExcelJs.Xlsx

open Fable.Core
open Fable.Core.JsInterop
open Fable.ExcelJs

type Xlsx =
    abstract member readFile: filename:string -> Async<unit>
    abstract member read: filename:System.IO.Stream -> Async<unit>