using System.Diagnostics;
using DosyaTipCevirici.Core;
using DosyaTipCevirici.Infrastructure;

namespace DosyaTipCevirici.WinForms;

public sealed class MainForm : Form
{
    private static readonly Color Ink = Color.FromArgb(29, 43, 66);
    private static readonly Color Muted = Color.FromArgb(103, 118, 139);
    private static readonly Color Accent = Color.FromArgb(35, 91, 184);
    private const int TargetColumn = 3;
    private const int StatusColumn = 4;

    private readonly ConversionService _service = new(new OfficeConversionEngine());
    private readonly List<string> _files = [];
    private readonly Dictionary<string, FileFormat> _targetsBySource = new(StringComparer.OrdinalIgnoreCase);
    private readonly ListView _fileList = new() { View = View.Details, FullRowSelect = true, Dock = DockStyle.Fill, HideSelection = false };
    private readonly ComboBox _sourceBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = nameof(FileFormat.Label), Dock = DockStyle.Fill };
    private readonly ComboBox _formatBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = nameof(FileFormat.Label), Dock = DockStyle.Fill };
    private readonly TextBox _outputBox = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Button _addButton = new();
    private readonly Button _removeButton = new();
    private readonly Button _clearButton = new();
    private readonly Button _browseButton = new();
    private readonly Button _convertButton = new();
    private readonly Button _cancelButton = new();
    private readonly ProgressBar _progress = new() { Dock = DockStyle.Fill };
    private readonly Label _status = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Label _note = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Label _count = new() { AutoSize = true };
    private readonly Label _emptyHint = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
    private bool _updatingSelectors;
    private CancellationTokenSource? _cancellation;

    public MainForm()
    {
        Text = "Dosya Tip Dönüştürücü";
        MinimumSize = new Size(860, 610);
        Size = new Size(980, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        BackColor = Color.FromArgb(244, 247, 251);
        AutoScaleMode = AutoScaleMode.Dpi;
        AllowDrop = true;
        _fileList.AllowDrop = true;
        _outputBox.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Dönüştürülen Dosyalar");
        var executableIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (executableIcon is not null) Icon = executableIcon;

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Margin = Padding.Empty, Padding = Padding.Empty };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 148));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
        Controls.Add(root);

        var header = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty, BackColor = Color.FromArgb(26, 52, 99), Padding = new Padding(26, 19, 26, 15) };
        var headerLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 66));
        headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        if (Icon is not null)
            headerLayout.Controls.Add(new PictureBox { Image = Icon.ToBitmap(), SizeMode = PictureBoxSizeMode.Zoom, Dock = DockStyle.Fill, Margin = new Padding(0, 2, 14, 2) }, 0, 0);
        var headerText = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        headerText.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
        headerText.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
        headerText.Controls.Add(new Label { Text = "Dosya Tip Dönüştürücü", ForeColor = Color.White, Font = new Font(Font.FontFamily, 20, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.BottomLeft }, 0, 0);
        headerText.Controls.Add(new Label { Text = "Dosyalarınızı kolayca istediğiniz türe dönüştürün", ForeColor = Color.FromArgb(211, 225, 247), Font = new Font(Font.FontFamily, 9), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        headerLayout.Controls.Add(headerText, 1, 0);
        header.Controls.Add(headerLayout);
        root.Controls.Add(header, 0, 0);

        var fileCard = CreateCard(new Padding(22, 20, 22, 8));
        var fileLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
        fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        fileLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
        var fileToolbar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
        fileToolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        fileToolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 392));
        var fileHeading = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        fileHeading.Controls.Add(new Label { Text = "Dosyalar", Font = new Font(Font.FontFamily, 12, FontStyle.Bold), ForeColor = Ink, AutoSize = true, Margin = new Padding(0, 8, 10, 0) });
        _count.ForeColor = Muted;
        _count.Margin = new Padding(0, 11, 0, 0);
        fileHeading.Controls.Add(_count);
        fileToolbar.Controls.Add(fileHeading, 0, 0);
        var fileActions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
        StyleButton(_addButton, "Dosya ekle", true, 110);
        StyleButton(_removeButton, "Seçileni kaldır", false, 136);
        StyleButton(_clearButton, "Listeyi temizle", false, 130);
        fileActions.Controls.AddRange([_addButton, _removeButton, _clearButton]);
        fileToolbar.Controls.Add(fileActions, 1, 0);
        fileLayout.Controls.Add(fileToolbar, 0, 0);

        _fileList.BorderStyle = BorderStyle.None;
        _fileList.BackColor = Color.White;
        _fileList.ForeColor = Ink;
        _fileList.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        _fileList.Columns.Add("Dosya");
        _fileList.Columns.Add("Tür");
        _fileList.Columns.Add("Boyut");
        _fileList.Columns.Add("Hedef");
        _fileList.Columns.Add("Durum");
        _fileList.Resize += (_, _) => ResizeFileColumns();
        var listHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        listHost.Controls.Add(_fileList);
        _emptyHint.Text = "Dosyaları buraya sürükleyin\nveya Dosya ekle düğmesini kullanın";
        _emptyHint.ForeColor = Muted;
        _emptyHint.Font = new Font(Font.FontFamily, 11);
        _emptyHint.AllowDrop = true;
        listHost.Controls.Add(_emptyHint);
        _emptyHint.BringToFront();
        fileLayout.Controls.Add(listHost, 0, 1);
        fileLayout.Controls.Add(new Label { Text = "Farklı türleri birlikte ekleyin; her kaynak türü için ayrı hedef seçin.", ForeColor = Muted, Font = new Font(Font.FontFamily, 9), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
        fileCard.Controls.Add(fileLayout);
        root.Controls.Add(fileCard, 0, 1);

        var settingsCard = CreateCard(new Padding(22, 8, 22, 18));
        var settings = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3 };
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        settings.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        settings.RowStyles.Add(new RowStyle(SizeType.Absolute, 41));
        settings.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        settings.Controls.Add(new Label { Text = "KAYNAK TÜR", ForeColor = Muted, Font = new Font(Font.FontFamily, 9, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        settings.Controls.Add(new Label { Text = "HEDEF TÜR", ForeColor = Muted, Font = new Font(Font.FontFamily, 9, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 1, 0);
        settings.Controls.Add(new Label { Text = "ÇIKTI KLASÖRÜ", ForeColor = Muted, Font = new Font(Font.FontFamily, 9, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 2, 0);
        _sourceBox.Margin = new Padding(0, 2, 14, 4);
        settings.Controls.Add(_sourceBox, 0, 1);
        _formatBox.Margin = new Padding(0, 2, 14, 4);
        settings.Controls.Add(_formatBox, 1, 1);
        var outputLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
        _outputBox.Margin = new Padding(0, 3, 8, 4);
        outputLayout.Controls.Add(_outputBox, 0, 0);
        StyleButton(_browseButton, "Gözat", false, 86);
        outputLayout.Controls.Add(_browseButton, 1, 0);
        settings.Controls.Add(outputLayout, 2, 1);
        _note.ForeColor = Muted;
        _note.Font = new Font(Font.FontFamily, 9);
        settings.Controls.Add(_note, 0, 2);
        settings.SetColumnSpan(_note, 3);
        settingsCard.Controls.Add(settings);
        root.Controls.Add(settingsCard, 0, 2);

        var footer = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(22, 12, 22, 12) };
        var footerLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
        footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240));
        footerLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        footerLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 7));
        _status.ForeColor = Ink;
        footerLayout.Controls.Add(_status, 0, 0);
        var finalActions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
        StyleButton(_convertButton, "Dönüştür", true, 126);
        StyleButton(_cancelButton, "İptal", false, 96);
        _cancelButton.Enabled = false;
        finalActions.Controls.AddRange([_convertButton, _cancelButton]);
        footerLayout.Controls.Add(finalActions, 1, 0);
        footerLayout.Controls.Add(_progress, 0, 1);
        footerLayout.SetColumnSpan(_progress, 2);
        footer.Controls.Add(footerLayout);
        root.Controls.Add(footer, 0, 3);

        _addButton.Click += (_, _) => AddFiles();
        _removeButton.Click += (_, _) => RemoveSelected();
        _clearButton.Click += (_, _) => { if (_cancellation is null) { _files.Clear(); _targetsBySource.Clear(); RefreshFiles(); } };
        _browseButton.Click += (_, _) => SelectOutput();
        _sourceBox.SelectedIndexChanged += (_, _) => { if (!_updatingSelectors) RefreshTargetChoices(); };
        _formatBox.SelectedIndexChanged += (_, _) => { if (!_updatingSelectors) ApplyTargetChoice(); };
        _convertButton.Click += async (_, _) => await ConvertFilesAsync();
        _cancelButton.Click += (_, _) => _cancellation?.Cancel();
        DragEnter += OnDragEnter;
        DragDrop += OnDragDrop;
        _fileList.DragEnter += OnDragEnter;
        _fileList.DragDrop += OnDragDrop;
        _fileList.SelectedIndexChanged += (_, _) =>
        {
            _removeButton.Enabled = _cancellation is null && _fileList.SelectedItems.Count > 0;
            if (_fileList.SelectedItems.Count == 0 || _updatingSelectors) return;
            var source = ConversionCatalog.Find((string)_fileList.SelectedItems[0].Tag!)!;
            for (var index = 0; index < _sourceBox.Items.Count; index++)
                if ((_sourceBox.Items[index] as FileFormat)?.Extension == source.Extension)
                    _sourceBox.SelectedIndex = index;
        };
        _emptyHint.DragEnter += OnDragEnter;
        _emptyHint.DragDrop += OnDragDrop;
        RefreshFiles();
    }

    private static Panel CreateCard(Padding margin)
    {
        var card = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Margin = margin, Padding = new Padding(18, 12, 18, 10) };
        card.Paint += (_, args) =>
        {
            using var pen = new Pen(Color.FromArgb(222, 229, 239));
            args.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
        };
        return card;
    }

    private static void StyleButton(Button button, string text, bool primary, int width)
    {
        button.Text = text;
        button.Width = width;
        button.Height = 34;
        button.Margin = new Padding(4, 3, 0, 0);
        button.Cursor = Cursors.Hand;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.FlatAppearance.BorderColor = Color.FromArgb(207, 217, 232);
        button.BackColor = primary ? Accent : Color.White;
        button.ForeColor = primary ? Color.White : Ink;
        button.Font = new Font("Segoe UI", 9, primary ? FontStyle.Bold : FontStyle.Regular);
    }

    private void ResizeFileColumns()
    {
        if (_fileList.Columns.Count != 5) return;
        var available = Math.Max(650, _fileList.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 8);
        _fileList.Columns[0].Width = Math.Max(235, (int)(available * 0.37));
        _fileList.Columns[1].Width = 74;
        _fileList.Columns[2].Width = 88;
        _fileList.Columns[3].Width = 76;
        _fileList.Columns[4].Width = Math.Max(160, available - _fileList.Columns[0].Width - 238);
    }

    private void AddFiles()
    {
        if (_cancellation is not null) return;
        var extensions = string.Join(';', ConversionCatalog.All.Select(format => "*." + format.Extension));
        using var dialog = new OpenFileDialog { Multiselect = true, Filter = $"Desteklenen dosyalar|{extensions}|Tüm dosyalar|*.*" };
        if (dialog.ShowDialog(this) == DialogResult.OK) AddPaths(dialog.FileNames);
    }

    private void AddPaths(IEnumerable<string> paths)
    {
        if (_cancellation is not null) return;
        foreach (var path in paths)
            if (File.Exists(path) && ConversionCatalog.Find(path) is not null && !_files.Contains(path, StringComparer.OrdinalIgnoreCase))
                _files.Add(path);
        RefreshFiles();
    }

    private void RemoveSelected()
    {
        if (_cancellation is not null) return;
        foreach (ListViewItem item in _fileList.SelectedItems)
            _files.Remove((string)item.Tag!);
        RefreshFiles();
    }

    private void SelectOutput()
    {
        if (_cancellation is not null) return;
        using var dialog = new FolderBrowserDialog { SelectedPath = Directory.Exists(_outputBox.Text) ? _outputBox.Text : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) };
        if (dialog.ShowDialog(this) == DialogResult.OK) _outputBox.Text = dialog.SelectedPath;
    }

    private void RefreshFiles()
    {
        var sourceFormats = _files.Select(path => ConversionCatalog.Find(path)!)
            .DistinctBy(format => format.Extension)
            .ToArray();
        var activeExtensions = sourceFormats.Select(format => format.Extension).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var stale in _targetsBySource.Keys.Where(key => !activeExtensions.Contains(key)).ToArray())
            _targetsBySource.Remove(stale);
        foreach (var source in sourceFormats)
            if (!_targetsBySource.TryGetValue(source.Extension, out var target) || !ConversionCatalog.CanConvert(source, target))
                _targetsBySource[source.Extension] = ConversionCatalog.TargetsFor(source).First();

        _updatingSelectors = true;
        var previouslySelected = (_sourceBox.SelectedItem as FileFormat)?.Extension;
        _sourceBox.Items.Clear();
        foreach (var source in sourceFormats) _sourceBox.Items.Add(source);
        var sourceIndex = Array.FindIndex(sourceFormats, format => format.Extension == previouslySelected);
        _sourceBox.SelectedIndex = sourceIndex >= 0 ? sourceIndex : sourceFormats.Length > 0 ? 0 : -1;
        _updatingSelectors = false;
        _sourceBox.Enabled = sourceFormats.Length > 0 && _cancellation is null;
        RefreshTargetChoices();

        _fileList.BeginUpdate();
        _fileList.Items.Clear();
        foreach (var path in _files)
        {
            var source = ConversionCatalog.Find(path)!;
            var item = new ListViewItem(Path.GetFileName(path)) { Tag = path, UseItemStyleForSubItems = false };
            item.SubItems.Add(source.Label);
            item.SubItems.Add(File.Exists(path) ? FormatSize(new FileInfo(path).Length) : "—");
            item.SubItems.Add(_targetsBySource[source.Extension].Label);
            item.SubItems.Add("Bekliyor");
            _fileList.Items.Add(item);
        }
        _fileList.EndUpdate();
        _count.Text = $"{_files.Count} dosya";
        _emptyHint.Visible = _files.Count == 0;
        _removeButton.Enabled = false;
        _clearButton.Enabled = _files.Count > 0 && _cancellation is null;
        ResizeFileColumns();

        _convertButton.Enabled = CanConvertAll() && _cancellation is null;
        _status.Text = _files.Count == 0 ? "Dosya ekleyin veya pencereye sürükleyin." : $"{_files.Count} dosya hazır.";
        UpdateNote();
    }

    private void RefreshTargetChoices()
    {
        _updatingSelectors = true;
        _formatBox.Items.Clear();
        if (_sourceBox.SelectedItem is FileFormat source)
        {
            foreach (var target in ConversionCatalog.TargetsFor(source)) _formatBox.Items.Add(target);
            var selected = _targetsBySource[source.Extension].Extension;
            var index = Enumerable.Range(0, _formatBox.Items.Count)
                .FirstOrDefault(item => (_formatBox.Items[item] as FileFormat)?.Extension == selected, -1);
            _formatBox.SelectedIndex = index;
        }
        _updatingSelectors = false;
        _formatBox.Enabled = _sourceBox.SelectedItem is not null && _cancellation is null;
        UpdateNote();
    }

    private void ApplyTargetChoice()
    {
        if (_sourceBox.SelectedItem is not FileFormat source || _formatBox.SelectedItem is not FileFormat target) return;
        _targetsBySource[source.Extension] = target;
        foreach (ListViewItem item in _fileList.Items)
            if (ConversionCatalog.Find((string)item.Tag!)?.Extension == source.Extension)
            {
                item.SubItems[TargetColumn].Text = target.Label;
                SetRowStatus(item, "Bekliyor", Ink);
            }
        _convertButton.Enabled = CanConvertAll() && _cancellation is null;
        UpdateNote();
    }

    private bool CanConvertAll() => _files.Count > 0 && _files.All(path =>
    {
        var source = ConversionCatalog.Find(path)!;
        return _targetsBySource.TryGetValue(source.Extension, out var target)
            && ConversionCatalog.CanConvert(source, target);
    });

    private static string FormatSize(long bytes) => bytes < 1024
        ? $"{bytes} B"
        : bytes < 1024 * 1024
            ? $"{bytes / 1024.0:0.#} KB"
            : $"{bytes / (1024.0 * 1024):0.#} MB";

    private static void SetRowStatus(ListViewItem item, string message, Color color)
    {
        item.SubItems[StatusColumn].Text = message;
        item.SubItems[StatusColumn].ForeColor = color;
    }

    private void UpdateNote()
    {
        var hasPdf = _files.Any(path => ConversionCatalog.Find(path)?.Family == FileFamily.Pdf);
        _note.Text = hasPdf
            ? "PDF dönüşümünde sayfa düzeni değişebilir. Her kaynak türü için ayrı hedef seçebilirsiniz."
            : "Seçilen hedef, aynı uzantıdaki tüm dosyalara uygulanır.";
    }

    private async Task ConvertFilesAsync()
    {
        if (_cancellation is not null || !CanConvertAll()) return;
        var jobs = _files.Select(path => (Path: path, Target: _targetsBySource[ConversionCatalog.Find(path)!.Extension])).ToArray();
        _cancellation = new CancellationTokenSource();
        _convertButton.Enabled = false;
        _cancelButton.Enabled = true;
        _addButton.Enabled = false;
        _removeButton.Enabled = false;
        _clearButton.Enabled = false;
        _browseButton.Enabled = false;
        _sourceBox.Enabled = false;
        _formatBox.Enabled = false;
        _progress.Maximum = jobs.Length;
        _progress.Value = 0;
        var succeeded = 0;
        var failed = 0;
        try
        {
            foreach (var job in jobs)
            {
                if (_cancellation.IsCancellationRequested) break;
                var item = _fileList.Items.Cast<ListViewItem>().First(row => Equals(row.Tag, job.Path));
                SetRowStatus(item, "Dönüştürülüyor...", Accent);
                _status.Text = Path.GetFileName(job.Path) + " işleniyor...";
                try
                {
                    var output = await _service.ConvertAsync(job.Path, _outputBox.Text, job.Target.Extension, _cancellation.Token);
                    SetRowStatus(item, "Tamamlandı: " + Path.GetFileName(output), Color.FromArgb(22, 128, 86));
                    succeeded++;
                }
                catch (OperationCanceledException) { SetRowStatus(item, "İptal edildi", Muted); break; }
                catch (Exception exception) { SetRowStatus(item, "Hata: " + exception.Message, Color.FromArgb(181, 55, 55)); failed++; }
                _progress.Value++;
            }
            _status.Text = $"Bitti: {succeeded} başarılı, {failed} hatalı.";
            if (succeeded > 0 && Directory.Exists(_outputBox.Text))
            {
                try { Process.Start(new ProcessStartInfo { FileName = _outputBox.Text, UseShellExecute = true }); }
                catch { /* Dönüştürme sonucunu etkilemez. */ }
            }
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;
            _cancelButton.Enabled = false;
            _addButton.Enabled = true;
            _clearButton.Enabled = _files.Count > 0;
            _removeButton.Enabled = _fileList.SelectedItems.Count > 0;
            _browseButton.Enabled = true;
            _sourceBox.Enabled = _sourceBox.Items.Count > 0;
            _formatBox.Enabled = _sourceBox.SelectedItem is not null;
            _convertButton.Enabled = CanConvertAll();
        }
    }

    private void OnDragEnter(object? sender, DragEventArgs args) =>
        args.Effect = _cancellation is null && args.Data?.GetDataPresent(DataFormats.FileDrop) == true
            ? DragDropEffects.Copy : DragDropEffects.None;

    private void OnDragDrop(object? sender, DragEventArgs args)
    {
        if (args.Data?.GetData(DataFormats.FileDrop) is string[] paths) AddPaths(paths);
    }
}
