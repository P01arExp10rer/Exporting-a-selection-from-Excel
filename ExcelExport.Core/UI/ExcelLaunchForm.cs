using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Exporting_a_selection_from_Excel.Excel;
using Exporting_a_selection_from_Excel.Workflow;

namespace Exporting_a_selection_from_Excel.UI
{
    public sealed class ExcelLaunchForm : Form
    {
        private TextBox _filePath;
        private Button _browse;
        private ComboBox _sheet;
        private TextBox _symbolColumn;
        private TextBox _symbols;
        private RadioButton _exact;
        private RadioButton _contains;
        private CheckBox _printToCmd;
        private CheckBox _placeMText;
        private NumericUpDown _maxRows;
        private Button _ok;
        private Button _cancel;

        public ExcelSearchSettings Settings { get; private set; }

        public ExcelLaunchForm()
        {
            Text = "Поиск в Excel → загрузка в чертеж";
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Width = 740;
            Height = 520;

            BuildUi();
        }

        private void BuildUi()
        {
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 9,
                Padding = new Padding(12),
                AutoSize = true,
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));

            table.Controls.Add(new Label { Text = "Файл Excel:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 0);
            _filePath = new TextBox { Dock = DockStyle.Fill };
            _filePath.TextChanged += (_, __) => TryReloadSheets();
            table.Controls.Add(_filePath, 1, 0);
            _browse = new Button { Text = "Обзор...", Dock = DockStyle.Fill };
            _browse.Click += (_, __) => BrowseExcel();
            table.Controls.Add(_browse, 2, 0);

            table.Controls.Add(new Label { Text = "Лист:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 1);
            _sheet = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDown };
            table.Controls.Add(_sheet, 1, 1);
            table.SetColumnSpan(_sheet, 2);

            table.Controls.Add(new Label { Text = "Столбец с символами:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 2);
            _symbolColumn = new TextBox { Dock = DockStyle.Fill };
            table.Controls.Add(_symbolColumn, 1, 2);
            table.SetColumnSpan(_symbolColumn, 2);

            table.Controls.Add(new Label { Text = "Символы для поиска:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 3);
            _symbols = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };
            _symbols.Height = 220;
            table.Controls.Add(_symbols, 1, 3);
            table.SetColumnSpan(_symbols, 2);

            table.Controls.Add(new Label { Text = "Режим совпадения:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 4);
            var modePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            _exact = new RadioButton { Text = "Точное", Checked = true, AutoSize = true };
            _contains = new RadioButton { Text = "Содержит", AutoSize = true };
            modePanel.Controls.Add(_exact);
            modePanel.Controls.Add(_contains);
            table.Controls.Add(modePanel, 1, 4);
            table.SetColumnSpan(modePanel, 2);

            table.Controls.Add(new Label { Text = "Вывод:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 5);
            var outPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            _printToCmd = new CheckBox { Text = "Печатать результаты в командной строке", Checked = true, AutoSize = true };
            _placeMText = new CheckBox { Text = "Разместить MText в чертеже", Checked = true, AutoSize = true };
            outPanel.Controls.Add(_printToCmd);
            outPanel.Controls.Add(_placeMText);
            table.Controls.Add(outPanel, 1, 5);
            table.SetColumnSpan(outPanel, 2);

            table.Controls.Add(new Label { Text = "Максимум строк:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 6);
            _maxRows = new NumericUpDown { Dock = DockStyle.Left, Minimum = 1, Maximum = 5000, Value = 200, Width = 120 };
            table.Controls.Add(_maxRows, 1, 6);
            table.SetColumnSpan(_maxRows, 2);

            var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true };
            _ok = new Button { Text = "OK", Width = 100 };
            _cancel = new Button { Text = "Отмена", Width = 100 };
            _ok.Click += (_, __) => OnOk();
            _cancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            buttonPanel.Controls.Add(_ok);
            buttonPanel.Controls.Add(_cancel);
            table.Controls.Add(buttonPanel, 0, 8);
            table.SetColumnSpan(buttonPanel, 3);

            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 240));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

            Controls.Add(table);
        }

        private void BrowseExcel()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Файлы Excel (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|Все файлы (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.Multiselect = false;
                if (!string.IsNullOrWhiteSpace(_filePath.Text) && File.Exists(_filePath.Text))
                    ofd.InitialDirectory = Path.GetDirectoryName(_filePath.Text);

                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    _filePath.Text = ofd.FileName;
                    TryReloadSheets();
                }
            }
        }

        private void TryReloadSheets()
        {
            var path = (_filePath.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            var current = _sheet.SelectedItem as string;
            var names = ExcelSheetNames.TryGetSheetNames(path);
            if (names.Count == 0)
                return;

            _sheet.BeginUpdate();
            try
            {
                _sheet.Items.Clear();
                foreach (var n in names) _sheet.Items.Add(n);
                if (!string.IsNullOrWhiteSpace(current) && names.Any(n => n.Equals(current, StringComparison.OrdinalIgnoreCase)))
                    _sheet.SelectedItem = names.First(n => n.Equals(current, StringComparison.OrdinalIgnoreCase));
                else
                    _sheet.SelectedIndex = 0;
            }
            finally
            {
                _sheet.EndUpdate();
            }
        }

        private void OnOk()
        {
            var settings = new ExcelSearchSettings
            {
                ExcelFilePath = _filePath.Text,
                SheetName = string.IsNullOrWhiteSpace(_sheet.Text) ? (_sheet.SelectedItem as string) : _sheet.Text,
                SymbolColumnName = _symbolColumn.Text,
                Symbols = SplitSymbols(_symbols.Text),
                MatchMode = _contains.Checked ? SymbolMatchMode.Contains : SymbolMatchMode.Exact,
                PrintToCommandLine = _printToCmd.Checked,
                PlaceMTextInDrawing = _placeMText.Checked,
                MaxRows = (int)_maxRows.Value,
            };

            settings.Normalize();

            var errors = new List<string>();
            if (string.IsNullOrWhiteSpace(settings.ExcelFilePath) || !File.Exists(settings.ExcelFilePath))
                errors.Add("Выберите существующий файл Excel.");
            if (string.IsNullOrWhiteSpace(settings.SheetName))
                errors.Add("Выберите лист.");
            if (string.IsNullOrWhiteSpace(settings.SymbolColumnName))
                errors.Add("Укажите имя столбца с символами (заголовок).");
            if (settings.Symbols == null || settings.Symbols.Count == 0)
                errors.Add("Введите хотя бы один символ (по одному в строке или через запятую/пробел).");

            if (errors.Count > 0)
            {
                MessageBox.Show(this, string.Join(Environment.NewLine, errors), "Не все поля заполнены", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Settings = settings;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static List<string> SplitSymbols(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return new List<string>();

            var seps = new[] { '\r', '\n', '\t', ',', ';', ' ' };
            return raw
                .Split(seps, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => (s ?? string.Empty).Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }
    }
}
