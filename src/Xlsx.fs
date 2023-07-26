[<AutoOpen>]
module Fable.ExcelJs.Xlsx

type Xlsx =
    /// read from a file
    abstract member readFile: filename:string -> Async<unit>
    /// read from a stream
    abstract member read: filename:System.IO.Stream -> Async<unit>
    /// load from a buffer
    abstract member load: filename:obj -> Async<unit>
    /// write to a file
    abstract member writeFile: filename:string -> Async<unit>
    /// write to a stream
    abstract member write: filename:System.IO.Stream -> Async<unit>
    /// write to a new buffer
    abstract member writeBuffer: unit -> Async<obj>