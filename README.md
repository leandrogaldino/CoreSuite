# CoreSuite

**A modular collection of reusable .NET 8 libraries, services, infrastructure components, and Windows Forms controls written in Visual Basic .NET.**

[![.NET](https://img.shields.io/badge/.NET-8.0-512BD4)](https://dotnet.microsoft.com/)
[![Visual Basic](https://img.shields.io/badge/Visual%20Basic-.NET-512BD4)](https://learn.microsoft.com/dotnet/visual-basic/)
[![Windows Forms](https://img.shields.io/badge/UI-Windows%20Forms-0078D4)](https://learn.microsoft.com/dotnet/desktop/winforms/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](./LICENSE)

## Overview

CoreSuite brings together independent packages created to reduce repeated application code and provide consistent building blocks for .NET 8 projects.

The solution includes:

- general-purpose helpers and extension methods;
- cryptography, file management, connectivity, database, and cloud services;
- Windows-specific infrastructure;
- reusable Windows Forms controls and designer components.

Each project is maintained as an independent NuGet package. Applications can install only the packages they need without referencing the complete solution.

> [!NOTE]
> This README is the catalog for the complete CoreSuite solution. Each project has its own README with installation instructions, examples, API details, behavior notes, and package-specific requirements.

## Projects

### Core libraries and infrastructure

These projects target plain `.NET 8` and can be used by desktop applications, services, console applications, workers, web applications, and other compatible projects.

| Project | Description |
|---|---|
| [CoreSuite.CodeCompiler](./CodeCompiler/) | Compiles Visual Basic and C# code at runtime through Roslyn, exposing references, diagnostics, invocation, and collectible assembly loading. |
| [CoreSuite.Connectivity](./Connectivity/) | Checks internet availability synchronously or asynchronously and monitors connection-state changes in the background. |
| [CoreSuite.Cryptography](./Cryptography/) | Provides SHA-256 hashing, PBKDF2 password hashing, and AES-256-GCM authenticated encryption. |
| [CoreSuite.ExceptionReporter](./ExceptionReporter/) | Captures structured exception reports and supports JSON serialization, atomic file storage, and asynchronous SMTP delivery. |
| [CoreSuite.Extensions](./Extensions/) | Provides reusable string transformations, word and whitespace utilities, and collection helper extensions. |
| [CoreSuite.FileManager](./FileManager/) | Copies and deletes files and directories asynchronously with progress reporting, cancellation, and exclusions. |
| [CoreSuite.FileMerger](./FileMerger/) | Creates and restores password-encrypted container files from directory trees, with metadata and progress reporting. |
| [CoreSuite.FileStateManager](./FileStateManager/) | Tracks pending file additions, replacements, and removals and applies them through transactional file operations. |
| [CoreSuite.FirebaseService](./FirebaseService/) | Accesses Firebase Authentication, Cloud Firestore, and Cloud Storage through a lightweight REST-based service. |
| [CoreSuite.Helpers](./Helpers/) | Centralizes common helpers for Brazilian formats, dates, enums, files, reflection, mathematics, text, and diagnostics. |
| [CoreSuite.JsonFileStore](./JsonFileStore/) | Persists strongly typed JSON data with atomic writes, backups, automatic recovery, and synchronous or asynchronous APIs. |
| [CoreSuite.Locator](./Locator/) | Provides a lightweight service locator with singleton, factory, keyed, and generic registrations. |
| [CoreSuite.MySqlService](./MySqlService/) | Executes MySQL queries, CRUD operations, stored procedures, and transactions and also provides database creation, backup, and restore. |

### Windows and Windows Forms infrastructure

These projects target `.NET 8 for Windows` and provide Windows-specific storage, helpers, extensions, and application services.

| Project | Description |
|---|---|
| [CoreSuite.AsyncLoader](./AsyncLoader/) | Displays a custom loading form over a parent form while asynchronous work is running and restores the original interface state afterward. |
| [CoreSuite.BusyOverlay](./BusyOverlay/) | Displays a cancellable busy overlay while synchronous or asynchronous operations execute, with controlled lifecycle and cleanup. |
| [CoreSuite.Cryptography.Windows](./Cryptography.Windows/) | Protects JSON objects and binary files through Windows DPAPI with user or machine scope and atomic persistence. |
| [CoreSuite.Extensions.Windows](./Extensions.Windows/) | Converts collections to `DataTable` instances and fills `DataGridView` controls while preserving relevant visual state. |
| [CoreSuite.Helpers.Windows](./Helpers.Windows/) | Provides Windows Forms helpers for control traversal, visibility, double buffering, and common image operations. |

### Windows Forms controls and components

These projects target `.NET 8 for Windows` and can be installed separately according to the interface requirements of each application.

| Project | Description |
|---|---|
| [CoreSuite.AnimatedBox](./AnimatedBox/) | Displays frame-based animations loaded from image collections or GIF files with configurable scaling. |
| [CoreSuite.AsyncLookupBox](./AsyncLookupBox/) | Performs asynchronous lookups and presents selectable results while retaining the selected value and related data. |
| [CoreSuite.CentralizedComboBox](./CentralizedComboBox/) | Centralizes and synchronizes reusable data sources and selection behavior across Windows Forms combo boxes. |
| [CoreSuite.CMessageBox](./CMessageBox/) | Provides customizable message dialogs with standard message types, exception details, localization, and error-reporting integration. |
| [CoreSuite.ColoredProgressBar](./ColoredProgressBar/) | Displays progress through a configurable range using customizable gradient colors. |
| [CoreSuite.ColorPicker](./ColorPicker/) | Embeds a Windows color-selection interface with palette, common, system, and custom color support. |
| [CoreSuite.ControlContainer](./ControlContainer/) | Hosts any Windows Forms control inside a floating drop-down anchored to another control. |
| [CoreSuite.ControlSlider](./ControlSlider/) | Adds a draggable divider for interactively resizing adjacent controls or interface regions. |
| [CoreSuite.CToolStripRender](./CToolStripRender/) | Applies reusable custom rendering and visual styling to `ToolStrip`, `MenuStrip`, and related controls. |
| [CoreSuite.CurrencyBox](./CurrencyBox/) | Provides culture-aware monetary input, formatting, validation, and strongly typed currency values. |
| [CoreSuite.DataGridViewContentCopy](./DataGridViewContentCopy/) | Adds configurable commands for copying individual cells or complete rows from a `DataGridView`. |
| [CoreSuite.DataGridViewFilterBox](./DataGridViewFilterBox/) | Filters a bound `DataGridView` through a reusable text-based interface and configurable column definitions. |
| [CoreSuite.DataGridViewLayoutManager](./DataGridViewLayoutManager/) | Saves and restores `DataGridView` column order, visibility, width, and related layout state. |
| [CoreSuite.DataGridViewNavigator](./DataGridViewNavigator/) | Provides record navigation and position controls for a `DataGridView` or its bound data source. |
| [CoreSuite.DateBox](./DateBox/) | Provides culture-aware masked date input with an integrated calendar drop-down. |
| [CoreSuite.DateTimeBox](./DateTimeBox/) | Combines culture-aware date and time input with a calendar, time selector, and explicit confirmation actions. |
| [CoreSuite.DateTimeBoxBase](./DateTimeBoxBase/) | Supplies shared culture, mask, parsing, formatting, value, and designer behavior for date and time controls. |
| [CoreSuite.DecimalBox](./DecimalBox/) | Provides culture-aware decimal input with configurable precision, grouping, and rounding. |
| [CoreSuite.FluidResizer](./FluidResizer/) | Resizes and repositions controls proportionally as their parent form or container changes size. |
| [CoreSuite.NavigationView](./NavigationView/) | Provides a configurable navigation pane that lazily creates, displays, caches, reloads, and disposes `UserControl` pages. |
| [CoreSuite.NoFocusCueButton](./NoFocusCueButton/) | Extends `Button` with built-in tooltip support while suppressing the standard dotted focus rectangle. |
| [CoreSuite.NumericBoxBase](./NumericBoxBase/) | Supplies shared parsing, formatting, culture, keyboard, and value behavior for numeric input controls. |
| [CoreSuite.PercentageBox](./PercentageBox/) | Provides culture-aware percentage input and formatting on top of the shared numeric infrastructure. |
| [CoreSuite.PictureViewer](./PictureViewer/) | Displays and navigates image collections with configurable navigation, inclusion, removal, and save actions. |
| [CoreSuite.QueriedBox](./QueriedBox/) | Performs database-backed lookups in a floating result grid while retaining the selected primary key and raw row values. |
| [CoreSuite.RichTextButton](./RichTextButton/) | Builds a button caption from individually styled text parts for richer visual emphasis. |
| [CoreSuite.Separator](./Separator/) | Provides a lightweight horizontal or vertical visual divider with configurable color and thickness. |
| [CoreSuite.SplitButton](./SplitButton/) | Combines a standard button action with a separate drop-down menu area. |
| [CoreSuite.TextBoxActionPanel](./TextBoxActionPanel/) | Attaches a configurable floating image-action panel to an existing `TextBoxBase` control. |
| [CoreSuite.TimeBox](./TimeBox/) | Provides culture-aware masked time input with an integrated time-selection drop-down. |
| [CoreSuite.ToolStripCheckBox](./ToolStripCheckBox/) | Hosts a standard check box inside a `ToolStrip`, `MenuStrip`, or compatible strip. |
| [CoreSuite.ValidationProvider](./ValidationProvider/) | Adds reusable control validation, custom validation logic, messages, and validation-state management to Windows Forms. |

## Installation

Every project is distributed independently. Install only the package required by the application:

```powershell
dotnet add package CoreSuite.Helpers
```

```powershell
dotnet add package CoreSuite.MySqlService
```

```powershell
dotnet add package CoreSuite.DateTimeBox
```

Replace the package name with the desired project from the tables above.

Dependencies between CoreSuite packages are declared by each project and are resolved automatically by NuGet.

## Using the source code

Clone the repository and restore the complete solution:

```powershell
git clone https://github.com/leandrogaldino/CoreSuite.git
cd CoreSuite
dotnet restore CoreSuite.sln
```

Build all projects in Release configuration:

```powershell
dotnet build CoreSuite.sln --configuration Release
```

Windows Forms projects require Windows and the .NET 8 SDK with Windows desktop support.

## Target frameworks

CoreSuite projects use one of the following target frameworks:

| Target | Intended use |
|---|---|
| `net8.0` | Cross-platform libraries, services, infrastructure, and data operations. |
| `net8.0-windows` | Windows-specific libraries and Windows Forms controls or components. |

## Package design

CoreSuite follows a modular package model:

- each production project has its own package, README, version, release notes, and XML documentation;
- projects can be consumed independently;
- shared behavior is extracted into focused base or infrastructure packages;
- Windows-specific code remains separate from general-purpose `.NET 8` libraries;
- package dependencies are explicit and installed transitively by NuGet.

## Documentation

Select a project in the catalog to open its directory and complete README.

Project-level documentation contains:

- package installation;
- namespaces and requirements;
- quick-start examples;
- public properties, methods, and events;
- designer usage where applicable;
- threading, disposal, security, and integration notes;
- known behaviors and limitations.

## Contributing

Contributions, issue reports, and improvement proposals are welcome.

Before submitting a change:

1. keep the project focused on its existing responsibility;
2. preserve compatibility with its declared target framework;
3. enable `Option Strict On` and `Option Explicit On` in Visual Basic projects;
4. document public APIs in English;
5. update the project README and package release notes when behavior changes;
6. build the complete solution before opening a pull request.

## License

CoreSuite is distributed under the [MIT License](./LICENSE).

## Author

Created and maintained by **Leandro Galdino**.
