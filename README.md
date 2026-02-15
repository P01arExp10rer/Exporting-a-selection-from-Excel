# ExcelExport.Gstar

Плагин для экспорта выборки из Excel в чертёж: поддерживаются **GstarCAD** и **AutoCAD 2024**.

## Структура решения

- **ExcelExport.Core** — общая логика (поиск в Excel, форма настроек, EPPlus). Без привязки к CAD.
- **ExcelExport.AutoCAD2024** — плагин для AutoCAD 2024 (разработка ведётся здесь).
- **ExcelExport.Gstar** — плагин для GstarCAD (.NET 8).

Изменения в Core автоматически попадают в оба плагина. Если меняете логику вставки в чертёж в AutoCAD — её нужно вручную перенести в GstarCAD (см. [docs/SYNC-AutoCAD-to-GstarCAD.md](docs/SYNC-AutoCAD-to-GstarCAD.md)).

## Зависимости (NuGet)

- **EPPlus** 6.1.2 — работа с Excel
- **GstarCADNET** 26.0.0 — [GstarCAD 2026 .NET API](https://www.nuget.org/packages/GstarCADNET/26.0.0/) (неофициальный пакет на NuGet)

Проект целевой платформы: **.NET 8.0** (Windows, WinForms).

## Сборка

```bash
dotnet restore ExcelExport.Gstar.sln
dotnet build ExcelExport.Gstar.sln -c Debug
```

Сборка создаёт библиотеку в `ExcelExport.Gstar\bin\Debug\net8.0-windows\ExcelExport.dll`. Подключайте её в GstarCAD как .NET-плагин (NETLOAD и т.п.).

## Отладка в CAD (AutoCAD 2024 или GstarCAD)

1. В Visual Studio выберите профиль запуска **«AutoCAD 2024»** или **«GstarCAD»** в выпадающем списке рядом с кнопкой запуска (или через «Проект» → «Свойства» → «Отладка»).
2. При необходимости измените путь к исполняемому файлу в `ExcelExport.Gstar\Properties\launchSettings.json` (если CAD установлен в другую папку).
3. Нажмите **F5**. Запустится выбранный CAD, к нему автоматически подключится отладчик.
4. В CAD выполните команду **NETLOAD** и укажите путь к `ExcelExport.dll` (папка `bin\Debug\net8.0-windows\` проекта).
5. Вызовите команду плагина (**ExcelExport** или **ExcelExportMenu**) — сработают точки останова.

### Сборка для AutoCAD 2024

В решении есть отдельный проект **ExcelExport.AutoCAD2024** ( .NET Framework 4.8, API Autodesk.AutoCAD ):

1. Убедитесь, что AutoCAD 2024 установлен (по умолчанию: `C:\Program Files\Autodesk\AutoCAD 2024`). Если путь другой, задайте при сборке:  
   `dotnet build ExcelExport.AutoCAD2024\ExcelExport.AutoCAD2024.csproj /p:AutoCAD2024Path=C:\ваш\путь`
2. Соберите проект:  
   `dotnet build ExcelExport.AutoCAD2024\ExcelExport.AutoCAD2024.csproj -c Debug`
3. В AutoCAD 2024: **NETLOAD** → укажите `ExcelExport.AutoCAD2024\bin\Debug\ExcelExport.dll`
4. Команды в AutoCAD: **ExcelExport** (сообщение о загрузке), **ExcelExportMenu** (окно поиска по Excel и вставка в чертёж).

Отладка: в Visual Studio выберите проект **ExcelExport.AutoCAD2024** как запускаемый, профиль «AutoCAD 2024» в списке запуска, затем F5. После запуска AutoCAD выполните NETLOAD и загрузите эту DLL — точки останова будут срабатывать.