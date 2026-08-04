# CoreSuite.Extensions

**General-purpose string and collection utilities for .NET 8 applications.**

> [!NOTE]
> `CoreSuite.Extensions` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8 and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.Extensions` provides small, reusable operations for manipulating strings and retrieving a single object from a collection. The package includes character and word reversal, word counting, diacritic removal, title and camel-case conversion, first-occurrence replacement, extra-space removal and object creation when a collection contains no matching element.

Because the package does not depend on Windows Forms, it can be used in desktop applications, services, console applications, class libraries and other .NET 8 projects.

## Features

- Reverses all characters in a string.
- Reverses the order of space-separated words.
- Counts words separated by spaces, tabs or line breaks.
- Removes accents and other Unicode non-spacing marks.
- Converts text to title case using the current culture.
- Converts space-, underscore- or hyphen-separated text to camel case.
- Replaces only the first occurrence of a substring.
- Collapses repeated spaces between text fragments.
- Retrieves a single collection element or creates a new instance when no element is found.
- Includes English XML documentation for the public API.
- Requires no third-party packages.

## Requirements

- .NET 8 (`net8.0`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Extensions
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Extensions
```

## Namespace

Import the namespace containing the extension modules:

```vb
Imports CoreSuite.Extensions.Extensions
```

## Quick start

```vb
Imports CoreSuite.Extensions.Extensions

Dim ReversedText As String = "CoreSuite".ReverseText()
Dim ReversedWords As String = "CoreSuite Extensions".ReverseWords()
Dim WordCount As Integer = "Reusable .NET utilities".CountWords()
Dim UnaccentedText As String = "Configuração".ToUnaccented()
Dim CamelName As String = "customer-order_number".ToCamel()
Dim CompactText As String = "CoreSuite    Extensions".RemoveExtraSpaces()
```

## API reference

### String extensions

| Member | Behavior |
| --- | --- |
| `ReverseText()` | Returns the characters in reverse order. |
| `ReverseWords()` | Returns space-separated words in reverse order. |
| `CountWords()` | Counts fragments separated by spaces, tabs, carriage returns or line feeds. |
| `ToUnaccented()` | Decomposes Unicode text and removes non-spacing marks. |
| `ToTitle()` | Converts text to title case using the current thread culture. |
| `ToCamel()` | Converts words separated by spaces, underscores or hyphens to camel case. |
| `ReplaceFirst(oldValue, newValue)` | Replaces only the first occurrence of `oldValue`. |
| `RemoveExtraSpaces()` | Removes empty fragments created by repeated space characters. |

### Collection utility

| Member | Behavior |
| --- | --- |
| `FirstOrNew(source, predicate)` | Returns the single matching element or attempts to create a new instance of the collection element type when no match exists. |

## String examples

### Reverse characters or words

```vb
Dim Characters As String = "Hello World".ReverseText()
Dim Words As String = "Hello World".ReverseWords()
```

The results are `dlroW olleH` and `World Hello`, respectively.

### Normalize text for comparison

```vb
Dim SearchText As String = "João da Silva".ToUnaccented().ToLowerInvariant()
```

`ToUnaccented` removes Unicode non-spacing marks. It does not perform case conversion, so a separate casing operation can be applied when required.

### Convert naming styles

```vb
Dim Title As String = "CUSTOMER ORDER".ToTitle()
Dim PropertyName As String = "customer-order_number".ToCamel()
```

`ToTitle` and the casing operations used by `ToCamel` are culture-sensitive.

### Replace only the first occurrence

```vb
Dim Result As String = "one two one".ReplaceFirst("one", "first")
```

The result is `first two one`.

### Remove repeated spaces

```vb
Dim Result As String = "CoreSuite    Extensions".RemoveExtraSpaces()
```

The result is `CoreSuite Extensions`. This method treats the space character as its separator; tabs and line breaks are not collapsed.

## Retrieve one element or create a new instance

`FirstOrNew` uses single-element semantics. It is useful when a query should return zero or one item and the calling code prefers a newly created object when no match exists.

```vb
Imports CoreSuite.Extensions.Extensions

Dim Customer As Customer = CollectionExtensions.FirstOrNew(Customers, Function(Item) Item.Id = CustomerId)
```

Important behavior:

- More than one matching element causes `InvalidOperationException`.
- When no element is found, the method attempts to create the item type through `Activator.CreateInstance`.
- `Nothing` is returned when the type cannot be created with a parameterless constructor.
- The source collection must not be `Nothing`.

## Empty-value behavior

| Member | `Nothing` or empty input |
| --- | --- |
| `ReverseText` | Returns `String.Empty`. |
| `ReverseWords` | Returns `String.Empty`. |
| `CountWords` | Returns `0`. |
| `ToUnaccented` | Returns `String.Empty`. |
| `ToTitle` | Returns `Nothing`. |
| `ToCamel` | Requires a non-`Nothing` string. |
| `ReplaceFirst` | Requires valid input and search values. |
| `RemoveExtraSpaces` | Requires a non-`Nothing` string. |

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Extensions` |
| Namespace | `CoreSuite.Extensions.Extensions` |
| Assembly | `CoreSuite.Extensions` |
| Target framework | `net8.0` |
| CoreSuite dependencies | None |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
