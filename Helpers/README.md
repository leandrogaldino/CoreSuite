# CoreSuite.Helpers

**A lightweight collection of reusable .NET 8 helpers for common application tasks, included in CoreSuite.**

> [!NOTE]
> `CoreSuite.Helpers` is one of the projects that make up the **CoreSuite** solution. This page documents the public API currently available in the Helpers library.

## Overview

CoreSuite.Helpers centralizes small, frequently repeated operations behind a consistent set of shared helper classes. It includes utilities for reflection, enumeration metadata, file-system operations, Brazilian data formats, Unix timestamps, SQL command inspection, internet connectivity checks, mathematical operations, and text generation.

The package targets plain `.NET 8` and does not depend on Windows Forms or third-party packages, making it suitable for desktop applications, services, console applications, class libraries, and other .NET projects.

## Key features

- Retrieve custom attributes from public shared or static members.
- Format and validate Brazilian CEP, phone, CPF, CNPJ, and email values.
- Print an expanded representation of a parameterized `DbCommand` for debugging.
- Convert between Unix epoch milliseconds and Brasília date/time.
- Read `DescriptionAttribute` metadata from enum values.
- Enumerate and filter enum members through reflection.
- Check whether files are locked and attempt safe file or directory deletion.
- Test internet connectivity synchronously or asynchronously.
- Find the nearest decimal value in a sample set.
- Evaluate simple, fully parenthesized arithmetic expressions.
- Inspect collection properties and invoke methods through reflection.
- Create runtime types with dynamically generated properties.
- Generate random strings and file names.
- Extract simple string values from compact JSON text.

## Requirements

- .NET 8 (`net8.0`)
- A reference to the `CoreSuite.Helpers` package, project, or assembly

The project uses only .NET base class library APIs and currently declares no external NuGet dependencies.

## Installation

### NuGet Package Manager

```powershell
Install-Package CoreSuite.Helpers
```

### .NET CLI

```bash
dotnet add package CoreSuite.Helpers
```

### Project reference

When using the CoreSuite source code directly, add the project reference below:

```xml
<ItemGroup>
  <ProjectReference Include="..\Helpers\Helpers.vbproj" />
</ItemGroup>
```

## Namespace

All public types are available through the following namespace:

```vb
Imports CoreSuite.Helpers
```

## Quick start

```vb
Imports CoreSuite.Helpers
Dim Cep As String = BrazilianFormatHelper.GetFormatedZipCode("12345678")
Dim IsCpfValid As Boolean = BrazilianFormatHelper.IsValidNaturalEntityDocument("52998224725")
Dim BrasiliaNow As Date = DateTimeHelper.Now()
Dim ClosestValue As Decimal = MathHelper.ApproximateValue({10D, 20D, 30D}, 24D)
Dim Token As String = TextHelper.GetRandomString(4, 4, "-", New List(Of CharFilter) From {CharFilter.UppercaseAlphabetic, CharFilter.Numeric})
```

## API summary

| Helper | Purpose |
|---|---|
| `AttributeHelper` | Retrieves custom attributes from public shared or static members. |
| `BrazilianFormatHelper` | Formats and validates Brazilian postal, phone, document, and email values. |
| `DatabaseHelper` | Prints a readable representation of a parameterized database command. |
| `DateTimeHelper` | Converts Unix milliseconds and retrieves the current Brasília date/time. |
| `EnumHelper` | Reads enum values, fields, and `DescriptionAttribute` metadata. |
| `FileHelper` | Checks file locks and attempts file-system deletions. |
| `InternetHelper` | Checks whether one of several public HTTP endpoints is reachable. |
| `MathHelper` | Finds approximate sample values and evaluates simple arithmetic expressions. |
| `ReflectionHelper` | Inspects collections, invokes methods, and creates dynamic runtime types. |
| `TextHelper` | Generates random text and extracts simple values from JSON-like text. |

## AttributeHelper

`AttributeHelper` retrieves a custom attribute from a named public shared or static member.

### `GetAttribute(Of TAttribute)`

```vb
Public Shared Function GetAttribute(Of TAttribute As Attribute)(Type As Type, MemberName As String) As TAttribute
```

The method searches with `BindingFlags.Public Or BindingFlags.Static`. It is particularly useful for enum fields and other public static members.

```vb
Imports System.ComponentModel
Imports CoreSuite.Helpers
Public Enum OrderStatus
    <Description("Waiting for payment")>
    Pending
    <Description("Completed")>
    Completed
End Enum
Dim Attribute As DescriptionAttribute = AttributeHelper.GetAttribute(Of DescriptionAttribute)(GetType(OrderStatus), NameOf(OrderStatus.Pending))
Dim Description As String = Attribute?.Description
```

The method returns `Nothing` when the member exists but does not contain the requested attribute. It throws an exception when the type or member name is invalid, or when the named member is not found.

## BrazilianFormatHelper

`BrazilianFormatHelper` contains formatting and validation methods for common Brazilian values.

> [!IMPORTANT]
> Formatting and validation are separate operations. `GetFormatedDocument` applies a CPF or CNPJ mask based on length, but it does not validate check digits. Use the corresponding validation method when authenticity must also be checked.

> [!NOTE]
> The public method names use the spelling `GetFormated...` with one `t`. This README preserves the current API names exactly.

### CEP

#### Format a CEP

```vb
Dim FormattedCep As String = BrazilianFormatHelper.GetFormatedZipCode("12345678")
```

Current output format:

```text
12.345-678
```

When the normalized value does not contain exactly eight digits, the original input is returned.

#### Validate a CEP

```vb
Dim IsValid As Boolean = BrazilianFormatHelper.IsValidZipCode("12345-678")
```

The method removes dots, hyphens, and spaces before checking for exactly eight numeric digits.

### Phone numbers

#### Detect the phone format

```vb
Dim Format As PhoneFormat = BrazilianFormatHelper.GetWhichPhoneFormat("11987654321")
```

Supported values:

| Value | Expected normalized structure |
|---|---|
| `FixedPhone` | Ten digits containing a supported two-digit DDD. |
| `CellPhone` | Eleven digits containing a supported DDD followed by `9`. |
| `SpecialPhone` | Eleven digits beginning with `0300`, `0500`, `0800`, or `0900`. |
| `InvalidPhone` | Any unsupported format. |

#### Format a phone number

```vb
Dim Mobile As String = BrazilianFormatHelper.GetFormatedPhoneNumber("11987654321")
Dim Landline As String = BrazilianFormatHelper.GetFormatedPhoneNumber("1134567890")
Dim Service As String = BrazilianFormatHelper.GetFormatedPhoneNumber("08001234567")
```

Typical results:

```text
(11) 9 8765-4321
(11) 3456-7890
0800-123-4567
```

Unsupported values are returned in normalized form, with parentheses, hyphens, and spaces removed.

### CPF and CNPJ

#### Apply a document mask

```vb
Dim Cpf As String = BrazilianFormatHelper.GetFormatedDocument("52998224725")
Dim Cnpj As String = BrazilianFormatHelper.GetFormatedDocument("11222333000181")
```

Typical results:

```text
529.982.247-25
11.222.333/0001-81
```

Only numeric values containing exactly 11 or 14 digits receive a mask.

#### Validate a CPF

```vb
Dim IsValidCpf As Boolean = BrazilianFormatHelper.IsValidNaturalEntityDocument("529.982.247-25")
```

The method normalizes the value, rejects known repeated-digit values, and validates both CPF check digits.

#### Validate a CNPJ

```vb
Dim IsValidCnpj As Boolean = BrazilianFormatHelper.IsValidLegalEntityDocument("11.222.333/0001-81")
```

The method normalizes the value and validates both CNPJ check digits.

### Email

```vb
Dim IsValidEmail As Boolean = BrazilianFormatHelper.IsValidEmail("user@example.com")
```

Email validation is performed through the regular expression embedded in the helper. It is intended as a structural application-level check and does not verify whether the mailbox or domain actually exists.

## DatabaseHelper

`DatabaseHelper` provides a debugging utility for ADO.NET commands.

### `DebugQuery`

```vb
Imports System.Data.Common
Imports CoreSuite.Helpers
Private Sub PrintCommand(Command As DbCommand)
    DatabaseHelper.DebugQuery(Command)
End Sub
```

The method starts with `Command.CommandText`, replaces each parameter name with its value surrounded by single quotes, and writes the result through `Debug.Print`.

> [!WARNING]
> The generated text is for diagnostics only. It is not a safe SQL serializer, does not escape values according to a database dialect, and must not be executed as a replacement for the original parameterized command.

## DateTimeHelper

`DateTimeHelper` converts Unix epoch milliseconds to and from the Brasília time zone and can return the current Brasília date/time independently of the machine's local time zone.

The implementation resolves the time zone using the identifier:

```text
E. South America Standard Time
```

### Convert Unix milliseconds to Brasília time

```vb
Dim BrasiliaDate As Date = DateTimeHelper.DateFromMilliseconds(0)
```

The input represents milliseconds elapsed since `1970-01-01 00:00:00 UTC`. The result is converted from UTC to the configured Brasília time zone.

### Convert Brasília time to Unix milliseconds

```vb
Dim Timestamp As Long = DateTimeHelper.MillisecondsFromDate(New Date(2026, 8, 1, 12, 0, 0))
```

The supplied `Date` is treated as an unspecified date/time belonging to the Brasília time zone and is converted to UTC before calculating the Unix timestamp.

### Get the current Brasília time

```vb
Dim CurrentDate As Date = DateTimeHelper.Now()
```

This method uses `DateTime.UtcNow` as its source and converts it to Brasília time.

> [!TIP]
> These methods are designed around Brasília time rather than arbitrary time zones. Use `DateTimeOffset` or direct `TimeZoneInfo` operations when an application needs to preserve offsets or support multiple zones.

## EnumHelper

`EnumHelper` simplifies access to enum members and their `DescriptionAttribute` metadata.

### Example enum

```vb
Imports System.ComponentModel
Public Enum PaymentStatus
    <Description("Waiting")>
    Pending
    <Description("Paid")>
    Paid
    Cancelled
End Enum
```

### Find a value by description

```vb
Dim Status As PaymentStatus = EnumHelper.GetEnumValue(Of PaymentStatus)("Paid")
```

The comparison is exact and case-sensitive. Only members containing `DescriptionAttribute` are considered. An exception is thrown when no matching description exists.

### Get all enum values

```vb
Dim Statuses As IEnumerable(Of PaymentStatus) = EnumHelper.GetEnumItems(Of PaymentStatus)()
```

### Filter enum values by field metadata

```vb
Dim Filtered As IEnumerable(Of PaymentStatus) = EnumHelper.GetEnumItems(Of PaymentStatus)(Function(Field) Field.Name <> NameOf(PaymentStatus.Cancelled))
```

The predicate receives the `FieldInfo` associated with each public static enum member.

### Get all descriptions

```vb
Dim Descriptions As IEnumerable(Of String) = EnumHelper.GetEnumDescriptions(Of PaymentStatus)()
```

A member without `DescriptionAttribute` contributes `String.Empty` to the returned sequence.

### Get one description

```vb
Dim Description As String = EnumHelper.GetEnumDescription(PaymentStatus.Pending)
```

The method returns `Nothing` when the selected enum member has no `DescriptionAttribute`.

## FileHelper

`FileHelper` contains small helpers for checking exclusive file access and attempting deletion operations.

### Check whether a file is locked

```vb
Imports System.IO
Dim File As New FileInfo("C:\Data\report.txt")
Dim IsLocked As Boolean = FileHelper.IsFileLocked(File)
```

The helper attempts to open the file for reading with `FileShare.None`. An `IOException` is interpreted as the file being locked or unavailable for exclusive access.

### Attempt to delete a file

```vb
Dim Deleted As Boolean = FileHelper.TryDeleteFile(New FileInfo("C:\Data\temporary.txt"))
```

The method attempts to acquire exclusive read/write access and then deletes the file. It returns `False` when an `IOException` occurs.

### Attempt to delete a directory recursively

```vb
Dim Deleted As Boolean = FileHelper.TryDeleteDirectory(New DirectoryInfo("C:\Data\Temporary"))
```

The directory and all its contents are deleted through `DirectoryInfo.Delete(True)`. The method returns `False` when an `IOException` occurs.

> [!IMPORTANT]
> These methods catch `IOException` only. Other failures, such as invalid arguments or some permission-related exceptions, can still propagate to the caller.

## InternetHelper

`InternetHelper` checks connectivity by contacting a predefined collection of public HTTPS endpoints. The current list includes Google, Microsoft, Facebook, AWS, Cloudflare, GitHub, and YouTube.

A shared `HttpClient` is used with a three-second timeout per request.

### Synchronous check

```vb
Dim IsOnline As Boolean = InternetHelper.IsInternetAvailable()
```

The synchronous method sends an HTTP `HEAD` request to each endpoint until one returns a successful status code.

> [!WARNING]
> This method blocks the calling thread while requests are running. Avoid calling it directly from a UI thread when temporary network delays would affect responsiveness.

### Asynchronous check

```vb
Private Async Function CheckConnectionAsync() As Task(Of Boolean)
    Return Await InternetHelper.IsInternetAvailableAsync()
End Function
```

The asynchronous method sends `GET` requests with `ResponseHeadersRead` and returns as soon as one endpoint responds successfully.

A `False` result means none of the configured endpoints returned a successful response. It does not necessarily prove that every network resource is unavailable; proxies, DNS settings, firewalls, endpoint blocking, or temporary remote failures can also affect the result.

## MathHelper

`MathHelper` provides nearest-value selection and a lightweight arithmetic evaluator.

### Find the closest sample value

```vb
Dim Samples As Decimal() = {5D, 10D, 25D, 50D}
Dim Result As Decimal = MathHelper.ApproximateValue(Samples, 18D)
```

The result is `25` because it has the smallest absolute difference from `18`. When two samples have the same difference, the first matching value is returned.

The sample array must contain at least one item.

### Evaluate a simple expression

```vb
Dim Result As Double = MathHelper.EvaluateExpression("((2*3)+4)")
```

The current evaluator processes one character at a time and is best treated as a minimal implementation for fully parenthesized expressions with single-digit operands and the operators `+`, `-`, `*`, and `/`.

> [!CAUTION]
> `EvaluateExpression` is not a general-purpose expression engine. Multi-digit numbers, decimal literals, unary operators, functions, variables, and normal operator-precedence parsing are not reliably supported by the current implementation.

## ReflectionHelper

`ReflectionHelper` contains utilities for inspecting collection properties, invoking methods, and creating types dynamically at runtime.

### Detect a collection property

```vb
Imports System.Reflection
Dim PropertyInfo As PropertyInfo = GetType(Customer).GetProperty(NameOf(Customer.Orders))
Dim IsCollection As Boolean = ReflectionHelper.IsCollection(PropertyInfo)
```

A property is considered a collection when its type implements `IEnumerable`. `String` is explicitly excluded.

### Get the element type of an indexed collection

```vb
Dim ItemType As Type = ReflectionHelper.GetCollectionPropertyType(PropertyInfo)
```

An overload accepts a `Type` directly:

```vb
Dim ItemType As Type = ReflectionHelper.GetCollectionPropertyType(GetType(List(Of Customer)))
```

The implementation locates an integer indexer named `get_Item`. It is intended for indexed collection types such as `List(Of T)` and can fail for non-indexed enumerables.

### Invoke a method

```vb
Imports System.Reflection
Dim Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic
Dim Result As Object = ReflectionHelper.InvokeMethod(Service, "Calculate", Flags, 10, 20)
```

The helper searches the runtime type and then its base types. Parameter matching uses the exact runtime type of each supplied argument. It returns `Nothing` when no matching method is found or when the invoked method has no return value.

> [!NOTE]
> A `Nothing` argument does not provide a runtime type for overload resolution and is not supported safely by the current matching logic.

### Create a runtime type

```vb
Dim PropertyNames As New List(Of String) From {"Id", "Name"}
Dim PropertyTypes As New List(Of Type) From {GetType(Integer), GetType(String)}
Dim RuntimeType As Type = ReflectionHelper.GetRunTimeType(PropertyNames, PropertyTypes)
Dim Instance As Object = Activator.CreateInstance(RuntimeType)
RuntimeType.GetProperty("Id").SetValue(Instance, 10)
RuntimeType.GetProperty("Name").SetValue(Instance, "Alice")
```

The generated type contains:

- a public parameterless constructor;
- one private backing field per property;
- public getter and setter methods;
- public properties using the supplied names and types.

The property-name and property-type lists must have the same number of elements.

## TextHelper

`TextHelper` provides random text generation, random file-name generation, and a lightweight string-based JSON value extractor.

### Generate a random string

```vb
Dim Filters As New List(Of CharFilter) From {
    CharFilter.UppercaseAlphabetic,
    CharFilter.Numeric
}
Dim Code As String = TextHelper.GetRandomString(3, 4, "-", Filters)
```

Example shape:

```text
A7KD-92QX-M4PL
```

Parameters:

| Parameter | Description |
|---|---|
| `SetCount` | Number of groups to generate. |
| `CharsPerSet` | Number of characters in each group. |
| `SetSeparator` | Text inserted between groups. |
| `PossibleChars` | Character categories used to build the available character set. |

Supported `CharFilter` values:

| Value | Characters |
|---|---|
| `Alphanumeric` | `0-9`, `A-Z`, and `a-z` |
| `Numeric` | `0-9` |
| `UppercaseAlphabetic` | `A-Z` |
| `LowercaseAlphabetic` | `a-z` |
| `SpecialCharacters` | Common punctuation and symbols |
| `Hexadecimal` | `0-9` and `A-F` |

Duplicate filter entries are removed before the character set is constructed.

> [!IMPORTANT]
> Although `PossibleChars` is declared optional, the current implementation calls it directly and therefore requires a non-`Nothing`, non-empty list. Pass at least one `CharFilter` value.

> [!WARNING]
> Random strings are generated with `System.Random` and are not cryptographically secure. Do not use this method for passwords, authentication tokens, encryption keys, or other security-sensitive values.

### Generate a random file name

```vb
Dim FileName As String = TextHelper.GetRandomFileName(".tmp")
```

The result combines:

- the current date and time using `ddMMyyyyHHmmss`;
- a GUID without separators;
- the random portion generated by `Path.GetRandomFileName`;
- the supplied extension text.

The extension is appended exactly as supplied, so include the leading dot when required.

### Extract a simple JSON string value

```vb
Dim Json As String = "{""name"":""Alice"",""status"":""active""}"
Dim Name As String = TextHelper.ExtractJsonValue(Json, "name")
```

The method returns `Alice` for the example above.

> [!CAUTION]
> This is a string-search helper, not a JSON parser. It expects the compact pattern `"key":"value"`, only extracts quoted string values, does not process escapes or nested structures, and can return `Nothing` when whitespace or formatting differs. Use `System.Text.Json` for general JSON processing.

## Error handling and input expectations

The helpers are intentionally small and generally expose the behavior of the underlying .NET APIs. Unless a method explicitly catches an exception or documents a fallback, callers should provide valid, non-`Nothing` arguments and handle relevant exceptions at the application boundary.

Common examples include:

- Brazilian formatting methods expect non-`Nothing` strings.
- `MathHelper.ApproximateValue` expects a non-empty sample array.
- Reflection helpers expect compatible types, members, and indexers.
- File helpers catch `IOException` but do not suppress every possible file-system exception.
- Connectivity checks suppress request failures and return `False` after all endpoints fail.

## Threading and performance

- Helper classes contain shared methods and do not require instances.
- `InternetHelper` reuses one shared `HttpClient`.
- `InternetHelper.IsInternetAvailable` performs blocking network I/O.
- `InternetHelper.IsInternetAvailableAsync` should be preferred in asynchronous and UI applications.
- `ReflectionHelper.GetRunTimeType` emits a new in-memory dynamic assembly and type on every call; cache the returned type when it will be reused.
- `DatabaseHelper.DebugQuery` is intended only for development diagnostics.

## Typical use cases

- Validating CPF or CNPJ values entered in a business application.
- Applying display masks to Brazilian contact and address data.
- Converting API timestamps to Brasília date/time.
- Populating user interfaces with enum values and descriptions.
- Inspecting a generated database command during development.
- Checking whether a temporary file can be deleted.
- Performing a lightweight online/offline check.
- Generating human-readable reference codes.
- Accessing non-public inherited methods in infrastructure code.
- Building temporary data shapes dynamically at runtime.

## Project information

- **Project:** Helpers
- **Assembly:** `CoreSuite.Helpers`
- **Namespace:** `CoreSuite.Helpers`
- **Framework:** `.NET 8`
- **External dependencies:** None
- **Repository:** Part of the CoreSuite solution

## License

See the license distributed with the CoreSuite repository for the terms that apply to this project.

CoreSuite.Helpers is designed to reduce repeated utility code while keeping common operations small, discoverable, and reusable across the CoreSuite ecosystem.
