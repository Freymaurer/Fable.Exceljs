module Workbook.Tests

open Fable.Mocha
open Fable.ExcelJs

let main = testList "Workbook" [
    testCase "Create Workbook" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        Expect.pass ()
    testCase "update properties 1" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let c = "Me"
        let d = "Lorem ipsum dolor med."
        let m = "Me2"
        let k = "key, word"
        let company = "CSB/DataPLAN"
        workbook.creator <- c
        workbook.description <- d 
        workbook.manager <- m
        workbook.keywords <- k
        workbook.company <- company
        Expect.equal workbook.creator c "c"
        Expect.equal workbook.description d "d"
        Expect.equal workbook.manager m "m"
        Expect.equal workbook.keywords k "k"
        Expect.equal workbook.company company "company"
    testCase "update properties 2" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let created = System.DateTime(2012,12,30)
        let modified = System.DateTime.Now
        let lastPrinted = System.DateTime.Now
        let category = "UnitTests"
        workbook.created <- created
        workbook.modified <- modified 
        workbook.lastPrinted <- lastPrinted
        workbook.category <- category
        Expect.equal workbook.created created "created"
        Expect.equal workbook.modified modified "modified"
        Expect.equal workbook.lastPrinted lastPrinted "lastPrinted"
        Expect.equal workbook.category category "category"
    testCase "Add worksheet" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let ws_name = "my_sheet1"
        let ws = workbook.addWorksheet(ws_name)
        Expect.equal ws.name ws_name "name"
    testCase "Add worksheet with props " <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let ws_name = "my_sheet1"
        let tabColor = {|argb="FF00FF00"|}
        let props = WorksheetProperties.create(
            tabColor= tabColor,
            outlineLevelCol = 12,
            outlineLevelRow = 13,
            defaultRowHeight = 14,
            defaultColWidth = 15,
            dyDescent = 16
        )
        let ws = workbook.addWorksheet(ws_name, props)
        Expect.equal ws.properties.tabColor tabColor "tabColor"
        Expect.equal ws.properties.outlineLevelCol 12 "outlineLevelCol"
        Expect.equal ws.properties.outlineLevelRow 13 "outlineLevelRow"
        Expect.equal ws.properties.defaultRowHeight 14 "defaultRowHeight"
        Expect.equal ws.properties.defaultColWidth 15 "defaultColWidth"
        Expect.equal ws.properties.dyDescent 16 "dyDescent"
    testCase "worksheets" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let ws_name = "my_sheet1"
        let ws = workbook.addWorksheet(ws_name)
        Expect.equal workbook.worksheets [|ws|] "worksheets"
    testCase "Remove worksheets" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let ws_name = "my_sheet1"
        let ws = workbook.addWorksheet(ws_name)
        Expect.equal workbook.worksheets [|ws|] "worksheets"
        workbook.removeWorksheet(ws.id)
        Expect.equal workbook.worksheets [||] "no worksheets"
    testCase "eachSheet" <| fun _ ->
        let workbook = ExcelJs.Excel.Workbook()
        let ws_name = "my_sheet1"
        let ws_name2 = "my_sheet2"
        let ws = workbook.addWorksheet(ws_name)
        let ws2 = workbook.addWorksheet(ws_name2)
        workbook.eachSheet(fun (ws, id) ->
            ws.name <- sprintf "New Name %i" id
        )
        Expect.equal ws.name "New Name 1" "name update 1"
        Expect.equal ws2.name "New Name 2" "name update 2"
]