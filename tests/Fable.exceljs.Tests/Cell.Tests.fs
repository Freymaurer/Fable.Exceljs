module Cell.Tests

open Fable.Mocha
open Fable.Core.JsInterop
open Fable.ExcelJs


[<Literal>]
let private ws_name = "MySheet1"

let private v_arr = [| Some <| box "TestEntry"; Some <| box "Test"; Some <| box 12; Some "Anything" |]

let private hyperlinkObj = {|
      text = "www.mylink.com"
      hyperlink = "http://www.mylink.com"
      tooltip = "www.mylink.com"
    |}

let private formulaObj = {| formula = "A2+A3" |}

let private richtTextObj = {| richText = [
        {|text = "This is"|}
        !!{|font = {|italic = true|}; text = "italic"|}
    ] |}

let private errorObj = {| error = "#VALUE!" |}

let private excelValueType_array = [|
    None
    Some 2
    Some !!2.12
    Some !!"Test me!"
    Some <| !!System.DateOnly(2023,4,14)
    Some <| !!hyperlinkObj
    Some <| !!formulaObj
    Some <| !!richtTextObj
    Some <| !!true
    Some <| !!errorObj
|]

let main = testList "Cell" [
    testCase "worksheet.getCell" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let cell_a1 = ws.getCell "A1"
        Expect.equal cell_a1.address "A1" "address"
        Expect.equal cell_a1.value None "value empty" 
        let row = ws.getRow(1)
        row.values <- v_arr
        let intermediate_cell_a1 = ws.getCell "A1"
        // This is strange but good to know behavior, reading cell will keep it like that even if values are updates through
        // another type (in this case row.values) ...
        Expect.equal cell_a1.value None "row values got replaced, old cell request" 
        // ... but reading again work just fine, the setter on the other hand ..
        Expect.equal intermediate_cell_a1.value (Some "TestEntry") "row values got replaced, cell read in new"
        // .. still works flawlessly on the cell with the outdated value
        cell_a1.value <- Some "newOption"
        Expect.equal cell_a1.value (Some "newOption") "value empty" 
    testCase "cell.type" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let row = ws.getRow(1)
        row.values <- !!excelValueType_array
        let a1 = ws.getCell "A1"
        Expect.equal a1.value None "a1 value"
        Expect.equal a1.``type`` 0 "a1 type"
        let b1 = ws.getCell "B1"
        Expect.equal b1.value (Some 2) "b1 value"
        Expect.equal b1.``type`` 2 "b1 type"
        let c1 = ws.getCell "C1"
        Expect.equal c1.value (Some 2.12) "c1 value"
        Expect.equal c1.``type`` 2 "c1 type"
        let d1 = ws.getCell "D1"
        Expect.equal d1.value (Some "Test me!") "d1 value"
        Expect.equal d1.``type`` 3 "d1 type"
        let e1 = ws.getCell "E1"
        Expect.equal e1.value (Some <| System.DateOnly(2023,4,14)) "e1 value"
        Expect.equal e1.``type`` 4 "e1 type"
        let f1 = ws.getCell "F1"
        Expect.equal f1.value (Some hyperlinkObj) "f1 value"
        Expect.equal f1.``type`` 5 "f1 type"
        let g1 = ws.getCell "G1"
        Expect.equal g1.value (Some formulaObj) "g1 value"
        Expect.equal g1.``type`` 6 "g1 type"
        let h1 = ws.getCell "H1"
        Expect.equal h1.value (Some richtTextObj) "h1 value"
        Expect.equal h1.``type`` 8 "h1 type"
        let i1 = ws.getCell "I1"
        Expect.equal i1.value (Some true) "i1 value"
        Expect.equal i1.``type`` 9 "i1 type"
        let j1 = ws.getCell "J1"
        Expect.equal j1.value (Some errorObj) "j1 value"
        Expect.equal j1.``type`` 10 "j1 type"
    testCase "cell.name" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        // A1 - D1
        let a1 = ws.getCell "A1"
        a1.name <- "TestCell"
        Expect.equal a1.name "TestCell" "name"
        a1.names <- [|"NextCell"; "BestCell";|]
        Expect.equal a1.name "NextCell" "name; is always first of names array"
        Expect.equal a1.names [|"NextCell"; "BestCell";|] "names"
        a1.removeName("NextCell")
        Expect.equal a1.name "BestCell" "name; is always first of names array 2"
        Expect.equal a1.names [|"BestCell";|] "names 2"
    ]