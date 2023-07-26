﻿module TestTasks

open BlackFox.Fake
open Fake.DotNet

open ProjectInfo
open BasicTasks

open ProjectInfo
open BasicTasks

[<Literal>]
let FableTestPath_input = "tests/Fable.exceljs.Tests"
[<Literal>]
let FableTestPath_output = "tests/Fable.exceljs.JsNativeTests/fable"

[<AutoOpen>]
module private Helper =

    open Fake
    open Fake.Core

    let createProcess exe arg dir =
        CreateProcess.fromRawCommandLine exe arg
        |> CreateProcess.withWorkingDirectory dir
        |> CreateProcess.ensureExitCode

    module Proc =

        module Parallel =

            open System

            let locker = obj()

            let colors = [|  
                ConsoleColor.DarkYellow
                ConsoleColor.DarkCyan 
                ConsoleColor.Magenta
                ConsoleColor.Blue
                ConsoleColor.Cyan
                ConsoleColor.DarkMagenta
                ConsoleColor.DarkBlue
                ConsoleColor.Yellow
            |]

            let print color (colored: string) (line: string) =
                lock locker
                    (fun () ->
                        let currentColor = Console.ForegroundColor
                        Console.ForegroundColor <- color
                        Console.Write colored
                        Console.ForegroundColor <- currentColor
                        Console.WriteLine line)

            let onStdout index name (line: string) =
                let color = colors.[index % colors.Length]
                if isNull line then
                    print color $"{name}: --- END ---" ""
                else if String.isNotNullOrEmpty line then
                    print color $"{name}: " line

            let onStderr name (line: string) =
                let color = ConsoleColor.Red
                if isNull line |> not then
                    print color $"{name}: " line

            let redirect (index, (name, createProcess)) =
                createProcess
                |> CreateProcess.redirectOutputIfNotRedirected
                |> CreateProcess.withOutputEvents (onStdout index name) (onStderr name)

            let printStarting indexed =
                for (index, (name, c: CreateProcess<_>)) in indexed do
                    let color = colors.[index % colors.Length]
                    let wd =
                        c.WorkingDirectory
                        |> Option.defaultValue ""
                    let exe = c.Command.Executable
                    let args = c.Command.Arguments.ToStartInfo
                    print color $"{name}: {wd}> {exe} {args}" ""

            let run cs =
                cs
                |> Seq.toArray
                |> Array.indexed
                |> fun x -> printStarting x; x
                |> Array.map redirect
                |> Array.Parallel.map Proc.run

    let dotnet = createProcess "dotnet"

    let npm =
        let npmPath =
            match ProcessUtils.tryFindFileOnPath "npm" with
            | Some path -> path
            | None ->
                "npm was not found in path. Please install it and make sure it's available from your path. " +
                "See https://safe-stack.github.io/docs/quickstart/#install-pre-requisites for more info"
                |> failwith

        createProcess npmPath

    let run proc arg dir =
        proc arg dir
        |> Proc.run
        |> ignore

    let runParallel processes =
        processes
        |> Proc.Parallel.run
        |> ignore
    
    let cleanFable = BuildTask.create "cleanFable" [clean; build] {
        System.IO.Directory.CreateDirectory FableTestPath_output |> ignore
        run dotnet "fable clean --yes" FableTestPath_output
    }

module RunTests = 

    /// runs `npm test` in root. 
    /// npm test consists of `test` and `pretest`
    /// check package.json in root for behavior
    let runTestsJs = BuildTask.create "runTestsJS" [clean; cleanFable; build] {
        run npm "test" ""
    }


let runTests = BuildTask.create "RunTests" [clean; build; RunTests.runTestsJs ] { 
    ()
}

module WatchTests =

    let private watchProjTests (projPath:string) =
        let pName = 
            let n = System.IO.Path.GetFileNameWithoutExtension(projPath)
            $"[{n}]"
        pName, dotnet "watch run" projPath

    let private fableTestsProcesses =
        [
            "[Fable Core]", dotnet $"fable watch {FableTestPath_input} -o {FableTestPath_output} --run npm run test:live" "."
            "[Mocha JsNative]", npm $"run testnative:live" "."
        ]

    let allTest = fableTestsProcesses

    let watchJS = BuildTask.create "watchTestsJS" [clean; build] {
        fableTestsProcesses
        |> runParallel
    }

let watchTests = BuildTask.create "watchTests" [clean; build] {
    WatchTests.allTest
    |> runParallel
}