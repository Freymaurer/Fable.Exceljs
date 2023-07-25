# Fable.exceljs

Fable bindings for [exceljs](https://github.com/exceljs/exceljs).

> Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

This is far from complete, but can be expanded as needed.

# install

## Femto

Using [Femto](https://github.com/Zaid-Ajaj/Femto), this will install both dotnet and npm dependency

```bash
femto install Fable.Exceljs
```

## Nuget and NPM

[Nuget](https://www.nuget.org/packages/Fable.Exceljs):

```bash
dotnet add package Fable.Exceljs --version 1.0.3
```

You must also install the correct [npm](https://www.npmjs.com/package/exceljs) dependency,
the current recommended version for `exceljs` is 4.3.0!

```bash
npm i exceljs
```

# develop

## prerequisites

- .NET 6 SDK
- nodejs (tested with ~v16)
- npm (tested with v9)

## setup

- dotnet tool restore
- npm install

## build

run `./build.cmd`

### or run the build project directly:

`dotnet run --project ./build/build.fsproj`

## test

- run `./build.cmd runTests`
- run `npm test`

As this repository only contains bindings for exceljs, the test will not run in dotnet but only in node.js environment.

## publish

1. increase version from latest release `./build.cmd releasenotes semver:xxx` [`semver:major`; `semver:minor`; `semver:patch`]
2. `./build.cmd pack`
3. upload package in `pkg`. 
