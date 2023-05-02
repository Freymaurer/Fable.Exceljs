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

#### or run the build project directly:

`dotnet run --project ./build/build.fsproj runTests`