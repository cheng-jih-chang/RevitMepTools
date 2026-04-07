# RevitAddinHotReloadDemo Architecture

## 1. Overall Architecture

This project uses a **Host + Logic separation architecture** so that a Revit Add-in can support a workflow similar to Hot Reload.

### Load Flow

```text
Revit
  ↓
RevitAddinHost.dll  (the stable host assembly)
  ↓
Loader.cs
  ↓
dist/RevitMepLogic.dll (the reloadable logic assembly)
```

Project structure:

```text
RevitAddinHotReloadDemo
│
├─ dist
│   RevitMepLogic.dll
│   RevitMepLogic.pdb
│
├─ RevitAddinHost
│   App.cs
│   CommandButton1.cs
│   CommandButton2.cs
│   Loader.cs
│   RevitAddinHost.csproj
│
└─ RevitMepLogic
    Entry.cs
    RevitMepLogic.csproj
    ```

---

## 2. Role of Each Project

### RevitAddinHost (Host)

This DLL is loaded directly by Revit.

Responsibilities:

- Create the Ribbon UI
- Handle button commands
- Load the Logic DLL

Main files:

App.cs
CommandButton1.cs
CommandButton2.cs
Loader.cs

When Revit starts:

Revit
 ↓
Load RevitAddinHost.dll
 ↓
App.OnStartup()
 ↓
Create Ribbon

---

### RevitMepLogic (Logic)

This is the **actual feature implementation** layer.

For example:

- Automatic annotation
- BOP calculation
- Sleeve generation
- Dynamo-to-C# converted features

Main entry point:

Entry.cs

Example:

namespace RevitMepLogic
{
    public class Entry
    {
        public string Run()
        {
            return "Hello from Logic";
        }
    }
}

---

## 3. What the Loader Does

The Loader loads `dist/RevitMepLogic.dll` at runtime.

byte[] asmBytes = File.ReadAllBytes(LogicDllPath);
Assembly asm = Assembly.Load(asmBytes);

Why this approach is used:

- Avoids Revit locking the DLL
- Allows the DLL in `dist` to be overwritten
- Supports a Hot Reload-like workflow

Execution flow:

```text
Revit
 ↓
CommandButton
 ↓
Loader.Call()
 ↓
Load dist/RevitMepLogic.dll
 ↓
Execute Entry.Run()
```

---

## 4. Why Use the `dist` Folder

The compiled DLL from RevitMepLogic is generated at:

RevitMepLogic/bin/Debug/RevitMepLogic.dll

But the DLL actually loaded by Revit is:

dist/RevitMepLogic.dll

So after building, the output must be copied into `dist`.

---

## 5. Daily Development Workflow (Hot Reload Style)

If you modify:

RevitMepLogic/Entry.cs

You only need to run:

dotnet build .\RevitMepLogic\RevitMepLogic.csproj

Copy-Item .\RevitMepLogic\bin\Debug\RevitMepLogic.dll .\dist\RevitMepLogic.dll -Force
Copy-Item .\RevitMepLogic\bin\Debug\RevitMepLogic.pdb .\dist\RevitMepLogic.pdb -Force

Then:

Go back to Revit
Click the Ribbon button

The updated Logic DLL will be loaded.

You do **not** need to:

- Restart Revit
- Rebuild the Host project

---

## 6. When Revit Must Be Restarted

Rule of thumb:

| Modified files       | Need to restart Revit? |
| -------------------- | ---------------------- |
| RevitMepLogic/*         | No                     |
| RevitAddinHost/*     | Yes                    |

Reason:

Revit loads:

RevitAddinHost.dll

during startup, and in .NET Framework:

Assemblies cannot be unloaded individually

So the Host DLL remains locked for the current Revit session.

---

## 7. Recommended Development Rhythm

Before the first launch:

dotnet build .\RevitMepAddinHost\RevitMepAddinHost.csproj

After that, during normal development, only build the Logic project.

dotnet build .\RevitMepLogic\RevitMepLogic.csproj

powershell -ExecutionPolicy Bypass -File .\build.ps1

powershell -ExecutionPolicy Bypass -File .\release.ps1

powershell -ExecutionPolicy Bypass -File .\release.ps1 -Version 0.1.0

---


## 8. Why Use the Host + Logic Architecture

The biggest limitation of Revit Add-ins is:

Revit does not support Hot Reload

The DLL is locked after being loaded.

The solution is:

Host (stable)
Logic (replaceable)

This allows the following workflow:

edit code
build
copy
click button

and the updated Logic can be used immediately.

---

## 9. Future Extensions

This architecture can be expanded into:

- Multi-command routing
- Automated build scripts
- DLL version checking
- Dynamic plugin systems

It is suitable for:

- Large Revit Add-ins
- Dynamo-to-C# migration projects
- BIM automation tools

---

## 10. Revit Add-in Registration Manifest (`.addin`)

Revit reads an XML manifest file at startup to know:

- which DLL to load
- which class is the entry point

Without this file, Revit does not know that your add-in exists.

To generate a new GUID in PowerShell:
```
[guid]::NewGuid().ToString().ToUpper()
```

Create this file at:

C:\Users\AppData\Roaming\Autodesk\Revit\Addins\2024\RevitAddinHotReloadDemo.addin

Example manifest:

```
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>RevitAddinHotReloadDemo</Name>
    <Assembly>D:\Project\BIM+C#\RevitAddinHotReloadDemo\RevitAddinHost\bin\Debug\RevitAddinHost.dll</Assembly>
    <AddInId>NEW-GUID-HERE</AddInId>
    <FullClassName>RevitAddinHost.App</FullClassName>
    <VendorId>TEST</VendorId>
    <VendorDescription>RevitAddinHotReloadDemo</VendorDescription>
  </AddIn>
</RevitAddIns>
```