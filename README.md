# VsSharp (VS#)

A lightweight toolkit for building Visual Studio extensions (VSIX). VsSharp wraps the
`EnvDTE` automation model and the VS Shell interop APIs behind a friendlier, strongly-typed
surface so you can focus on your extension's behavior instead of boilerplate.

Published on NuGet as [`VsSharp`](https://www.nuget.org/packages/VsSharp).

## Features

- **Host** — a singleton entry point that wires up DTE events (solution, document, startup,
  shutdown) and exposes convenience accessors (`Solution`, `Windows`, `Statusbar`, output
  window helpers, service resolution).
- **Command framework** — declare buttons, menus and context menus in XML and have them
  built on the VS command bars at runtime, including nested sub-menus and dynamically
  generated menus.
- **Dependency-driven enabling** — show/enable commands based on context such as whether a
  solution, project, or document is open.
- **Project & solution extensions** — a rich set of `EnvDTE.Project` / `Solution` extension
  methods for references, project items, properties, package metadata, and more.
- **Document & window accessors** — read/write the active document's text and manage tool
  windows.

## Requirements

- Visual Studio (extension development workload)
- `net472`
- The VS SDK is referenced via the `Microsoft.VisualStudio.SDK` metapackage

## Installation

```powershell
Install-Package VsSharp
```

or via the .NET CLI:

```bash
dotnet add package VsSharp
```

## Getting started

### 1. Initialize the Host

Set up the `Host` from your package's initialization. Assigning `Host.Instance.DTE`
automatically subscribes to the relevant DTE events.

```csharp
using CnSharp.VisualStudio.Extensions;
using Microsoft.VisualStudio.Shell;

protected override async Task InitializeAsync(CancellationToken ct, IProgress<ServiceProgressData> progress)
{
    await this.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

    var dte = (EnvDTE._DTE)await GetServiceAsync(typeof(EnvDTE._DTE));
    Host.Instance.DTE = dte;

    // optional lifecycle hooks
    Host.Instance.SolutionOpenedAction = () => { /* ... */ };
    Host.Instance.DocumentOpenedAction = doc => { /* ... */ };
}
```

### 2. Define commands in XML

Commands are described by a `CommandConfig` and typically loaded from an embedded XML
resource. Each control points at a class implementing `ICommand` via its `ClassName`.

```xml
<root>
  <resource>MyExtension.Resources.Strings</resource>
  <buttons>
    <button id="MyExtension.SayHello"
            class="MyExtension.Commands.SayHelloCommand"
            dependsOn="SolutionProject" />
  </buttons>
  <menus>
    <menu id="MyExtension.TopMenu" caption="My Tools">
      <menu id="MyExtension.Child" class="MyExtension.Commands.ChildCommand" />
    </menu>
  </menus>
</root>
```

### 3. Implement a command

```csharp
using CnSharp.VisualStudio.Extensions.Commands;

public class SayHelloCommand : ICommand
{
    public void Execute(object args = null)
    {
        Host.Instance.WriteToOutputWindow("MyExtension", "Hello from VsSharp!");
    }
}
```

### 4. Load the command manager

```csharp
var plugin = new Plugin
{
    Id = new Guid("..."),
    Name = "MyExtension",
    Assembly = typeof(MyPackage).Assembly,
    CommandConfig = config // deserialized from the XML above
};

var manager = new CommandManager(plugin);
manager.Load();        // builds the command bar controls
// manager.Execute("MyExtension.SayHello");
// manager.Disconnect(); // tears the controls down
```

Commands declared with a `dependsOn` value (for example `SolutionProject` or `Document`)
are automatically enabled or disabled as the corresponding context changes, via
`CommandManager.ApplyDependencies`.

## Working with projects

`ProjectExtensions` adds many helpers on top of `EnvDTE.Project`:

```csharp
string ns       = project.GetRootNameSpace();
string dir      = project.GetDirectory();
bool isSdkStyle = project.IsSdkBased();
bool isCore     = project.IsNetCoreProject();

// references
project.AddReference("Newtonsoft.Json", @"C:\libs\Newtonsoft.Json.dll");
var refs = project.GetReferences();

// project items
project.AddFromFile(@"C:\path\file.cs");
project.AddFromFileAsLink(@"C:\shared\Common.cs", "Common.cs");
project.AddFolder("Generated");

// package metadata (NuGet PropertyGroup values)
PackageProjectProperties meta = project.GetPackageProjectProperties();
meta.Description = "Updated description";
List<string> failures = project.SavePackageProjectProperties(meta);
```

## Key types

| Type | Purpose |
| --- | --- |
| [`Host`](src/CnSharp.VisualStudio.Extensions/Host.cs) | Singleton entry point; DTE event wiring and shell convenience accessors. |
| [`Plugin`](src/CnSharp.VisualStudio.Extensions/Plugin.cs) | Describes an extension instance (id, assembly, command config, resources). |
| [`CommandManager`](src/CnSharp.VisualStudio.Extensions/Commands/CommandManager.cs) | Builds, resets, executes and tears down command bar controls. |
| [`CommandConfig`](src/CnSharp.VisualStudio.Extensions/Commands/CommandConfig.cs) / [`CommandMenu`](src/CnSharp.VisualStudio.Extensions/Commands/CommandMenu.cs) / [`CommandButton`](src/CnSharp.VisualStudio.Extensions/Commands/CommandButton.cs) | XML-serializable command definitions. |
| [`ICommand`](src/CnSharp.VisualStudio.Extensions/Commands/ICommand.cs) | Implemented by command handler classes. |
| [`ICommandMenuGenerator`](src/CnSharp.VisualStudio.Extensions/Commands/ICommandMenuGenerator.cs) | Produces dynamically generated sub-menus. |
| [`ProjectExtensions`](src/CnSharp.VisualStudio.Extensions/Extensions/ProjectExtensions.cs) / [`SolutionExtensions`](src/CnSharp.VisualStudio.Extensions/Extensions/SolutionExtensions.cs) | `EnvDTE` automation helpers. |
| [`IDocumentAccessor`](src/CnSharp.VisualStudio.Extensions/IDocumentAccessor.cs) / [`IWindowAccessor`](src/CnSharp.VisualStudio.Extensions/IWindowAccessor.cs) / [`IToolWindowAccessor`](src/CnSharp.VisualStudio.Extensions/IToolWindowAccessor.cs) | Document and window access. |
| [`PackageProjectProperties`](src/CnSharp.VisualStudio.Extensions/Projects/PackageProjectProperties.cs) / [`ProjectProperties`](src/CnSharp.VisualStudio.Extensions/Projects/ProjectProperties.cs) | Strongly-typed project/package metadata. |

## License

MIT © CnSharp Studio. See the repository for details.

Project home: <https://github.com/cnsharp/VS.Sharp>
