# Project

## What we are building

`CofyExcelParser` is a C# Excel parser used by CatSweeper's Info pipeline to turn tabular data into the intermediate `DataTable`/`DataRow` model that becomes generated JSON.

## Stack

| Layer | Choice | Notes |
|---|---|---|
| Language | C# | .NET 8 |
| Consumer | CatSweeper Info pipeline | Imported as Git subtree |
| Project | `CofyExcelParser.csproj` | Build with `dotnet build` |

## Domain in one paragraph

The parser reads Excel/CSV source files and produces a generic table model. CatSweeper's `InfoBuilder` then converts that table into `data.json` files consumed by runtime Info classes. Changes to parsing semantics affect which data rows survive into the game.

## Non-obvious constraints

- Checked out inside CatSweeper as a Git subtree at `Modules/CofyExcelParser/`. Only the programmer/owner pushes changes back to `https://github.com/felixwongong/CofyExcelParser.git` via `Tools/subtree.ps1`.
- Any breaking change must be coordinated with `dotnet build CatSweeper.sln` and CatSweeper Info pipeline tests.

## CatSweeper usage

See CatSweeper's `.agent/systems/info-pipeline.md` for the full pipeline.
