# Перенос изменений из AutoCAD в GstarCAD

Разработка ведётся в проекте **ExcelExport.AutoCAD2024**. Общая логика (Excel, форма, настройки) вынесена в **ExcelExport.Core**, поэтому перенос в GstarCAD в большинстве случаев автоматический.

## Что переносится само

- **ExcelExport.Core** — общая библиотека, на которую ссылаются оба CAD-проекта.
- Любые изменения в папках **Excel**, **UI**, **Workflow\ExcelSearchSettings** делайте **только в ExcelExport.Core**.
- После этого достаточно пересобрать решение: и AutoCAD, и GstarCAD получат обновлённую логику без копирования кода.

## Что нужно синхронизировать вручную

Логика **вставки в чертёж** (MText, запрос точки и т.п.) разная у каждой CAD-платформы и лежит в:

- **ExcelExport.AutoCAD2024:** `Workflow\ExcelToDrawingLoader.cs` (API Autodesk.AutoCAD)
- **GstarCAD:** `ExcelExport.Gstar\Workflow\ExcelToDrawingLoader.cs` (API Gssoft.Gscad)

Если вы меняете в AutoCAD-версии, например:

- способ формирования текста (BuildPlainText, ToMTextContents),
- ограничение по строкам, приглашение ввода точки,
- любую другую логику «что и как вставлять в чертёж»,

то те же изменения нужно внести в **GstarCAD-версию** `ExcelToDrawingLoader.cs`, подставив эквиваленты API GstarCAD (Document, Editor, MText, Transaction и т.д.).

## Краткий чек-лист

| Где правили | Действие для GstarCAD |
|-------------|------------------------|
| Excel, UI, настройки (Core) | Ничего — пересобрать решение достаточно. |
| ExcelExportCommands (команды) | При изменении имён/поведения команд — поправить `ExcelExport.Gstar\Class1.cs` (CommandMethod, вызовы GscadApp). |
| ExcelToDrawingLoader (вставка в чертёж) | Вручную перенести логику в `ExcelExport.Gstar\Workflow\ExcelToDrawingLoader.cs`, заменив вызовы AutoCAD API на GstarCAD. |

## Сборка GstarCAD после правок

```bash
dotnet build "ExcelExport.Gstar.sln" -c Debug
```

Или в Visual Studio: собрать решение (F6). Плагин для GstarCAD: `ExcelExport.Gstar\bin\Debug\net8.0-windows\ExcelExport.dll`.
