# CoreSuite Locator

**Lightweight singleton, factory and keyed service resolution for .NET 8 applications.**

> [!NOTE]
> `CoreSuite.Locator` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8 and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.Locator` provides a small, process-wide service registry for applications that need simple dependency resolution without configuring a full dependency injection framework. Services can be registered as shared singleton instances or through factories that create a new instance whenever the service is requested.

Each registration can also have a string key, allowing several implementations of the same service type to coexist. Attempting to resolve an unknown type and key combination raises a dedicated `ServiceNotRegisteredException`.

Because the package does not depend on Windows Forms, it can be used in desktop applications, services, console applications, class libraries and other .NET 8 projects.

## Features

- Registers an existing object as a singleton.
- Registers a factory that creates a new instance for every resolution.
- Resolves services through strongly typed generic methods.
- Supports multiple registrations of the same type through optional string keys.
- Replaces an existing registration when the same type and key are registered again.
- Reports missing registrations through `ServiceNotRegisteredException`.
- Uses a process-wide registry accessible through shared members.
- Includes English XML documentation for the public API.
- Requires no CoreSuite or third-party packages.

## Requirements

- .NET 8 (`net8.0`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Locator
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Locator
```

## Namespace

Import the namespace containing the locator and its exception type:

```vb
Imports CoreSuite.Infraestructure
```

## Quick start

Register an instance once and resolve it wherever the service is required:

```vb
Imports CoreSuite.Infraestructure

Dim Settings As New ApplicationSettings With {
    .ApplicationName = "CoreSuite Sample"
}

Locator.RegisterSingleton(Of ApplicationSettings)(Settings)

Dim CurrentSettings As ApplicationSettings = Locator.GetInstance(Of ApplicationSettings)()
```

Every resolution of a singleton registration returns the same object.

## Registration methods

| Member | Behavior |
| --- | --- |
| `RegisterSingleton(Of T)(instance, key)` | Stores an existing instance and returns it for every matching resolution. |
| `RegisterFactory(Of T)(factory, key)` | Stores a factory and invokes it for every matching resolution. |
| `GetInstance(Of T)(key)` | Resolves the registration matching the requested type and key. |

The optional `key` argument defaults to an empty string. A registration is uniquely identified by the combination of its service type and key.

## Register a singleton

Singleton registration is useful for shared state or services whose lifetime is managed by the application:

```vb
Dim Cache As New ApplicationCache()

Locator.RegisterSingleton(Of IApplicationCache)(Cache)

Dim FirstReference As IApplicationCache = Locator.GetInstance(Of IApplicationCache)()
Dim SecondReference As IApplicationCache = Locator.GetInstance(Of IApplicationCache)()
```

`FirstReference` and `SecondReference` refer to the same instance.

The generic type determines how the service is registered. Registering a concrete object through an interface makes it resolvable through that interface:

```vb
Locator.RegisterSingleton(Of IApplicationCache)(New ApplicationCache())
```

## Register a factory

Factory registration is useful when each consumer should receive a new object:

```vb
Locator.RegisterFactory(Of OperationContext)(Function() New OperationContext())

Dim FirstContext As OperationContext = Locator.GetInstance(Of OperationContext)()
Dim SecondContext As OperationContext = Locator.GetInstance(Of OperationContext)()
```

The factory is invoked by each call to `GetInstance`, so `FirstContext` and `SecondContext` are different instances. Any exception raised by the factory is propagated to the caller.

## Use keyed registrations

Keys allow different implementations or configurations to be registered for the same service type:

```vb
Locator.RegisterSingleton(Of IMessageSender)(New EmailMessageSender(), "email")
Locator.RegisterSingleton(Of IMessageSender)(New SmsMessageSender(), "sms")

Dim EmailSender As IMessageSender = Locator.GetInstance(Of IMessageSender)("email")
Dim SmsSender As IMessageSender = Locator.GetInstance(Of IMessageSender)("sms")
```

Keys are matched exactly. The default unkeyed registration and a keyed registration are independent entries.

## Replace a registration

Registering the same service type with the same key replaces the previous registration:

```vb
Locator.RegisterSingleton(Of IMessageSender)(New DevelopmentMessageSender(), "default")
Locator.RegisterSingleton(Of IMessageSender)(New ProductionMessageSender(), "default")

Dim Sender As IMessageSender = Locator.GetInstance(Of IMessageSender)("default")
```

`Sender` resolves to the most recently registered instance.

## Missing services

`GetInstance` throws `ServiceNotRegisteredException` when the requested type and key combination has not been registered:

```vb
Try
    Dim Sender As IMessageSender = Locator.GetInstance(Of IMessageSender)("push")
Catch Ex As ServiceNotRegisteredException
    Console.WriteLine(Ex.Message)
End Try
```

The exception derives from `InvalidOperationException` and identifies the unresolved service type and, when supplied, its key.

## Lifetime and ownership

`Locator` deliberately provides only singleton and factory registrations:

| Registration | Instance behavior | Ownership |
| --- | --- | --- |
| Singleton | Returns the registered object every time. | The application owns and disposes the object when appropriate. |
| Factory | Creates an object for every resolution. | The caller owns each returned object and disposes it when appropriate. |

The locator does not automatically dispose services or provide scoped lifetimes.

## Usage notes

- Registrations are stored globally for the lifetime of the process.
- There is no built-in unregister or clear operation.
- Register all services during application startup before resolving them concurrently.
- The registry does not synchronize writes; avoid modifying registrations from multiple threads at the same time.
- Factories do not receive the locator or automatically resolve constructor dependencies.
- Prefer constructor injection when explicit dependencies and managed lifetimes are required; use the locator for small applications or integration points where direct resolution is appropriate.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Locator` |
| Namespace | `CoreSuite.Infraestructure` |
| Assembly | `CoreSuite.Locator` |
| Target framework | `net8.0` |
| CoreSuite dependencies | None |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
