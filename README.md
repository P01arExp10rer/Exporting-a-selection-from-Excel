# Exporting a selection from Excel

Плагин для GstarCAD: экспорт выборки из Excel в чертёж.

## Зависимости (NuGet)

- **EPPlus** 6.1.2 — работа с Excel
- **GstarCADNET** 26.0.0 — [GstarCAD 2026 .NET API](https://www.nuget.org/packages/GstarCADNET/26.0.0/) (неофициальный пакет на NuGet)

Проект целевой платформы: **.NET 8.0** (Windows, WinForms).

## Сборка

```bash
dotnet restore Exporting-a-selection-from-Excel.sln
dotnet build Exporting-a-selection-from-Excel.sln -c Debug
```

Сборка создаёт библиотеку в `Exporting a selection from Excel\bin\Debug\net8.0-windows\`. Подключайте её в GstarCAD как .NET-плагин (NETLOAD и т.п.).