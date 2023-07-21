# Fable.exceljs

Fable bindings for [exceljs](https://github.com/exceljs/exceljs).

> Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

This is far from complete, but can be expanded as needed.

# develop

### prerequisites

- .NET 6 SDK
- nodejs (tested with ~v16)
- npm (tested with v9)

### setup

- dotnet tool restore
- npm install

### build

#### Windows

run `build.cmd`

#### or run the build project directly:

`dotnet run --project ./build/build.fsproj`

### test

#### Windows

- run `build.cmd runTests`
- run `npm test`

As this repository only contains bindings for exceljs, the test will not run in dotnet but only in node.js environment.
