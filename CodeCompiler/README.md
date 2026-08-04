# CoreSuite.CodeCompiler

**Compile and execute Visual Basic or C# source code at runtime with Roslyn.**

> [!NOTE]
> `CoreSuite.CodeCompiler` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and uses the official Roslyn Visual Basic and C# compiler packages.

## Overview

`CoreSuite.CodeCompiler` turns Visual Basic or C# source text into an in-memory .NET assembly. The result can be inspected as a normal `Assembly`, or its public shared, static and instance methods can be invoked through a simple API.

Platform assemblies and physical assemblies already loaded by the host application are added as compilation references automatically. Additional references can be supplied by file path, loaded `Assembly` or `Type`.

Each generated assembly is loaded into a collectible `AssemblyLoadContext`. Disposing the returned `CompiledCode` releases the package's references to that context and requests unloading.

## Features

- Compiles Visual Basic and C# source code at runtime.
- Uses the latest language version supported by the included Roslyn packages.
- Compiles generated assemblies in release optimization mode.
- Enables `Option Strict On` for dynamically compiled Visual Basic code.
- Automatically references .NET platform assemblies.
- Automatically references suitable assemblies already loaded by the application.
- Adds custom references by path, `Assembly` or `Type`.
- Compiles source supplied as text or read from `.vb` and `.cs` files.
- Invokes public shared, static or instance methods.
- Creates an instance automatically for public instance methods.
- Resolves compatible overloads and performs selected argument conversions.
- Returns detailed compiler diagnostics through `CodeCompilationException`.
- Loads generated code in a collectible assembly load context.
- Includes English XML documentation for the public API.

## Requirements

- .NET 8 (`net8.0`)
- `Microsoft.CodeAnalysis.VisualBasic` 5.6.0, installed automatically by NuGet
- `Microsoft.CodeAnalysis.CSharp` 5.6.0, installed automatically by NuGet

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.CodeCompiler
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.CodeCompiler
```

## Namespace

Import the service namespace:

```vb
Imports CoreSuite.Services
```

## Quick start

The shortest path is `Execute`, which compiles the source, invokes one public method and disposes the generated assembly automatically:

```vb
Imports CoreSuite.Services

Dim sourceCode As String =
    "Namespace RuntimeSamples" & Environment.NewLine &
    "    Public NotInheritable Class Calculator" & Environment.NewLine &
    "        Public Shared Function Add(firstValue As Integer, secondValue As Integer) As Integer" & Environment.NewLine &
    "            Return firstValue + secondValue" & Environment.NewLine &
    "        End Function" & Environment.NewLine &
    "    End Class" & Environment.NewLine &
    "End Namespace"

Dim compiler As New CodeCompiler()
Dim result As Integer = CInt(compiler.Execute(sourceCode, CodeLanguage.VisualBasic, "RuntimeSamples.Calculator", "Add", 12, 8))
```

`result` is `20`.

> [!WARNING]
> Dynamically compiled code runs with the permissions of the host process. This package is not a sandbox. Do not compile or execute source supplied by untrusted users.

## Compile without immediate execution

Use `Compile` when several methods must be called, the generated assembly must be inspected, or its lifetime must extend beyond one invocation:

```vb
Dim compiler As New CodeCompiler()

Using compiledCode As CompiledCode = compiler.Compile(sourceCode, CodeLanguage.VisualBasic)
    Dim firstResult As Integer = CInt(compiledCode.Invoke("RuntimeSamples.Calculator", "Add", 2, 3))
    Dim secondResult As Integer = CInt(compiledCode.Invoke("RuntimeSamples.Calculator", "Add", 10, 5))
    Dim generatedAssembly As Reflection.Assembly = compiledCode.Assembly
End Using
```

Always dispose `CompiledCode`. A `Using` block is the simplest way to do so.

### Custom assembly name

When `assemblyName` is omitted, the compiler generates a unique name beginning with `DynamicCode_`. Supply a name when the generated assembly must be identifiable:

```vb
Using compiledCode As CompiledCode = compiler.Compile(sourceCode, CodeLanguage.VisualBasic, "RuntimeCalculations")
    Console.WriteLine(compiledCode.Assembly.GetName().Name)
End Using
```

## Compile C#

Pass `CodeLanguage.CSharp` for C# source:

```vb
Dim sourceCode As String =
    "namespace RuntimeSamples" & Environment.NewLine &
    "{" & Environment.NewLine &
    "    public static class TextTools" & Environment.NewLine &
    "    {" & Environment.NewLine &
    "        public static string Repeat(string value, int count)" & Environment.NewLine &
    "        {" & Environment.NewLine &
    "            return string.Concat(Enumerable.Repeat(value, count));" & Environment.NewLine &
    "        }" & Environment.NewLine &
    "    }" & Environment.NewLine &
    "}"

Dim compiler As New CodeCompiler()
Dim result As String = CStr(compiler.Execute(sourceCode, CodeLanguage.CSharp, "RuntimeSamples.TextTools", "Repeat", "Core", 3))
```

The result is `CoreCoreCore`.

## Compile source files

`CompileFile` detects the language from `.vb` and `.cs` extensions:

```vb
Using compiledCode As CompiledCode = compiler.CompileFile("C:\RuntimeCode\ReportFormula.vb")
    Dim result As Object = compiledCode.Invoke("ReportFormula", "Calculate", 150D, 0.1D)
End Using
```

Only `.vb` and `.cs` are recognized automatically. A different extension can still be compiled by explicitly specifying the language:

```vb
Using compiledCode As CompiledCode = compiler.CompileFile("C:\RuntimeCode\formula.txt", CodeLanguage.VisualBasic)
    Dim result As Object = compiledCode.Invoke("Formula", "Calculate", 100D)
End Using
```

Use `ExecuteFile` when only one method call is required:

```vb
Dim result As Object = compiler.ExecuteFile("C:\RuntimeCode\Rules.cs", "Rules.InvoiceRules", "Validate", invoiceTotal)
```

## Add compilation references

The compiler automatically adds platform assemblies and suitable assemblies already loaded in the current process. Add an explicit reference when the dynamic source uses another library that has not been loaded or must be resolved from a particular file.

### Reference an assembly path

```vb
compiler.AddReference("C:\Libraries\BusinessRules.dll")
```

The file must be a valid physical .NET assembly.

### Reference a loaded assembly

```vb
compiler.AddReference(GetType(MyBusinessService).Assembly)
```

### Reference the assembly containing a type

```vb
compiler.AddReference(GetType(MyBusinessService))
```

The three overloads add both the metadata reference required for compilation and, when necessary, information used to resolve that dependency when the generated assembly runs.

## Method invocation rules

`CompiledCode.Invoke` searches for a compatible public method using the exact type and method names supplied.

| Rule | Behavior |
| --- | --- |
| Type name | Must be the full runtime type name, including its namespace. |
| Visibility | Only public methods are considered. |
| Shared/static methods | Invoked without creating an object. |
| Instance methods | The declaring type must be non-abstract and have a public parameterless constructor. |
| Parameter count | Must exactly match the supplied argument count. Optional parameters are not filled automatically. |
| Generic methods | Open generic methods are ignored. |
| `ByRef` parameters | Not supported by the invocation helper. |
| Overloads | The most directly compatible overload is selected; equally suitable overloads cause `AmbiguousMatchException`. |
| Exceptions | Exceptions thrown by invoked code are rethrown as their original inner exception. |

The invocation helper accepts exact types, assignable types and selected conversions involving primitive `IConvertible` values, enums, nullable types and `Guid` values supplied as strings.

### Invoke an instance method

```vb
Dim sourceCode As String =
    "Public Class Greeter" & Environment.NewLine &
    "    Public Function CreateMessage(name As String) As String" & Environment.NewLine &
    "        Return ""Hello, "" & name & ""!""" & Environment.NewLine &
    "    End Function" & Environment.NewLine &
    "End Class"

Dim message As String = CStr(compiler.Execute(sourceCode, CodeLanguage.VisualBasic, "Greeter", "CreateMessage", "Leandro"))
```

`Greeter` is instantiated automatically through its public parameterless constructor.

## Handle compilation errors

Invalid source causes `CodeCompilationException`. Its `Diagnostics` property contains the Roslyn error messages:

```vb
Try
    Using compiledCode As CompiledCode = compiler.Compile(sourceCode, CodeLanguage.VisualBasic)
    End Using
Catch ex As CodeCompilationException
    For Each diagnostic As String In ex.Diagnostics
        Console.WriteLine(diagnostic)
    Next
End Try
```

Warnings do not prevent compilation; the exception contains errors that caused emission to fail.

## Public types

| Type | Purpose |
| --- | --- |
| `CodeCompiler` | Manages references and compiles or executes source. |
| `CompiledCode` | Owns a generated assembly and invokes its public methods. Implements `IDisposable`. |
| `CodeLanguage` | Selects `VisualBasic` or `CSharp`. |
| `CodeCompilationException` | Reports failed compilation and exposes diagnostic messages. |

## `CodeCompiler` API

| Member | Description |
| --- | --- |
| `AddReference(assemblyPath)` | Adds a physical assembly file. |
| `AddReference(referenceAssembly)` | Adds a loaded assembly with a physical location. |
| `AddReference(referenceType)` | Adds the assembly that defines a type. |
| `Compile(sourceCode, language, assemblyName)` | Compiles text and returns a disposable generated assembly. |
| `Execute(sourceCode, language, typeName, methodName, parameters)` | Compiles text, invokes one method and disposes the assembly. |
| `CompileFile(filePath)` | Compiles a `.vb` or `.cs` file using extension detection. |
| `CompileFile(filePath, language)` | Compiles a file using an explicit language. |
| `ExecuteFile(filePath, typeName, methodName, parameters)` | Detects the language, invokes one method and disposes the assembly. |
| `ExecuteFile(filePath, language, typeName, methodName, parameters)` | Uses an explicit language, invokes one method and disposes the assembly. |

## Lifetime and unloading

`CompiledCode.Dispose` clears its assembly reference and calls `Unload` on the collectible load context. Actual unloading is finalized by the .NET garbage collector after no application references remain to the generated assembly, its types, instances or delegates.

Do not store `CompiledCode.Assembly`, generated `Type` objects or generated instances longer than necessary if unloadability matters.

## Usage notes

- Validate or author source code before compiling it; this API provides execution, not isolation.
- Prefer a long-lived `CodeCompiler` when the same custom references are reused.
- Dispose every `CompiledCode` result.
- Use the full type name when invoking a method declared inside a namespace.
- Source file compilation reads the entire file as text.
- Generated output is an in-memory library, not an executable file on disk.
- The compiler currently emits one syntax tree per compilation request.

## License

This package is licensed under the MIT License.
