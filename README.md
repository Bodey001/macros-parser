# VBA Macro Parser

A C# console application that parses VBA source files and exports the structured macro data to **JSON**, **XML**, **plain text**, and **CSV** — with no third-party dependencies.

---

## Prerequisites

| Requirement | Version |
|---|---|
| .NET SDK | 10.0 or later |

Download the SDK: <https://aka.ms/dotnet/download>

---

## Project Structure

```
c#-macros-parser/
├── input/                            # Drop your VBA source file here
├── output/                           # Generated output subfolders appear here
│   └── {filename}-parsed/
│       ├── {filename}.json
│       ├── {filename}.xml
│       ├── {filename}.txt
│       └── {filename}.csv
├── src/
│   └── VbaMacroParser/               # Main console application
│       ├── Models/                   # Parse result data model
│       ├── Parser/                   # VbaLexer + VbaParser
│       ├── Exporters/                # JSON, XML, Text, CSV exporters
│       └── IO/                       # FileReader, OutputManager
└── tests/
    └── VbaMacroParser.Tests/         # MSTest unit tests
```

---

## Supported Input Formats

| Extension | Description |
|---|---|
| `.bas` | Standard VBA module |
| `.cls` | Class module |
| `.frm` | UserForm module |
| `.vba` | Raw VBA code |
| `.txt` | Plain text containing VBA code |
| `.vb` | Visual Basic source file |

> **Binary Office formats** (`.xlsm`, `.docm`, `.xlsb`) are **not supported** without a
> compound-file-binary (CFB) library. Extract the VBA source first and pass the resulting `.bas` / `.cls` files.

---

## Step-by-Step Usage

### 1. Clone / open the repository

```powershell
git clone https://github.com/Bodey001/macros-parser
cd macros-parser
```

### 2. Place your VBA source file in `input/`

```powershell
Copy-Item "C:\path\to\MyMacros.bas" input\
```

Any text-based VBA file will work.  If you exported macros from Excel or Word via the VBA editor
(**File → Export File**) you get `.bas`, `.cls`, or `.frm` files that are accepted directly.

### 3. Build and Test the project

Simply run the included setup script. It will ensure you have the .NET 10 SDK installed, build the project, and run the test suite:

```powershell
.\setup.bat
```

### 4. Run (auto-detect from `input/`)

You can run the parser using the included run script:

```powershell
.\run.bat
```

The application will automatically find and process **all** VBA files placed in the `input/` folder.

### 5. Run with an explicit file path

```powershell
dotnet run --project src\VbaMacroParser -- input\MyMacros.bas
```

You can pass any absolute or relative path — the file does not have to live in `input/`.

### 6. Inspect the outputs

A subfolder is created under `output/` named after your file:

```
output\MyMacros-parsed\
    MyMacros.json    ← full structured parse result
    MyMacros.xml     ← same data as well-formed XML
    MyMacros.txt     ← human-readable report
    MyMacros.csv     ← one row per procedure
```

---

## Running Unit Tests

```powershell
dotnet test tests\VbaMacroParser.Tests
```

The test suite covers:

| File | What is tested |
|---|---|
| `ParserTests.cs` | Sub/Function/Property parsing, parameters, variables, constants, comments, edge cases |
| `ExporterTests.cs` | JSON validity, XML well-formedness, CSV structure, text report content |
| `FileReaderTests.cs` | Encoding detection (UTF-8 BOM, UTF-16 LE, ANSI), unsupported format rejection |

---

## What Gets Parsed

For each VBA module the parser extracts:

- **Module name and type** (Standard / Class / Form) — detected from `Attribute VB_Name` and file extension
- **Option declarations** — `Option Explicit`, `Option Base 1`, etc.
- **Module-level comments** — `'` apostrophe and `Rem` style
- **Constants** — `Public/Private Const Name As Type = Value`
- **Variables** — `Dim`, `Public`, `Private`, `Friend`, `Global`, `Static` — including arrays and array bounds
- **Procedures** — `Sub`, `Function`, `Property Get/Let/Set`
  - Scope (`Public`, `Private`, `Friend`, or default)
  - `Static` flag
  - Return type (Functions and Property Get)
  - Parameters — name, type, `ByRef`/`ByVal`, `Optional`, `ParamArray`, default values
  - Line start / line end
  - Procedure body (raw text)
  - Inline and trailing comments

---

## Output Format Details

### JSON
Uses `System.Text.Json.JsonSerializer` (in-box since .NET 5). Fields use camelCase names.
The root object follows this shape:

```json
{
  "sourceFile": "MyMacros.bas",
  "parsedAt": "2026-04-17T12:00:00Z",
  "parserVersion": "1.0.0",
  "modules": [
    {
      "name": "MyMacros",
      "type": "Standard",
      "options": ["Explicit"],
      "constants": [...],
      "variables": [...],
      "procedures": [...]
    }
  ]
}
```

### XML
Produced via `System.Xml.XmlWriter`. Root element is `<VbaParseResult>`.

### Plain Text
Human-readable report with module headings, procedure signatures, parameter tables, and comment listings.

### CSV
RFC-4180 compliant. One row per procedure. Columns:

```
Module, ModuleType, ProcedureName, Kind, Scope, IsStatic, ReturnType,
Parameters, LineStart, LineEnd, CommentCount, BodyLineCount
```

---

## Architecture Overview

```
FileReader          reads and detects encoding
    ↓
VbaLexer            tokenises lines → Token stream
    ↓
VbaParser           state machine → VbaParseResult
    ↓
OutputManager       orchestrates exporters
    ├── JsonExporter   → {name}.json
    ├── XmlExporter    → {name}.xml
    ├── TextExporter   → {name}.txt
    └── CsvExporter    → {name}.csv
```

All types are in the `VbaMacroParser` namespace. No NuGet packages beyond MSTest are used.
