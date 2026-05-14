using System.ComponentModel;
using System.Runtime.InteropServices;

namespace ExcelReader;

public partial class MainForm : Form, IReaderHost
{
    private IDocumentReader? _currentReader;
    private ExcelDocumentReader? _excelReader;
    private string _lastRowColStats = string.Empty;

    private const int WM_SETREDRAW = 0x0B;

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    public MainForm()
    {
        InitializeComponent();
        SetupDataGridStyle();
    }

    // ═══════════════════ IReaderHost implementation ═══════════════════

    public void ShowExcelView()
    {
        wordContainer.Visible = false;
        dataGrid.Visible = true;
        bool hasSheets = cmbSheet.Items.Count > 0;
        lblSheetLabel.Visible = hasSheets;
        cmbSheet.Visible = hasSheets;
    }

    public void ShowWordPdfView()
    {
        dataGrid.Visible = false;
        wordContainer.Visible = true;
        lblSheetLabel.Visible = false;
        cmbSheet.Visible = false;
    }

    public void BeginUpdate()
    {
        dataGrid.SuspendLayout();
        if (wordView.IsHandleCreated)
            SendMessage(wordView.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
    }

    public void EndUpdate()
    {
        if (wordView.IsHandleCreated)
        {
            SendMessage(wordView.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            wordView.Invalidate();
        }
        dataGrid.ResumeLayout(true);
    }

    public void SetSheets(IReadOnlyList<string> sheetNames)
    {
        cmbSheet.Items.Clear();
        foreach (var name in sheetNames)
            cmbSheet.Items.Add(name);

        bool hasSheets = cmbSheet.Items.Count > 0;
        cmbSheet.Enabled = hasSheets;
        lblSheetLabel.Visible = hasSheets && dataGrid.Visible;
        cmbSheet.Visible = hasSheets && dataGrid.Visible;
        if (hasSheets)
            cmbSheet.SelectedIndex = 0;
    }

    string IReaderHost.SelectedSheet => cmbSheet.SelectedItem as string ?? string.Empty;

    public void ClearDataGrid()
    {
        dataGrid.Columns.Clear();
        dataGrid.Rows.Clear();
    }

    public void AddDataGridColumn(string name, string headerText)
    {
        var col = new DataGridViewTextBoxColumn
        {
            Name = name,
            HeaderText = headerText,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                WrapMode = DataGridViewTriState.False
            }
        };
        dataGrid.Columns.Add(col);
    }

    public void AddDataGridRow(object[] values)
    {
        dataGrid.Rows.Add(values);
    }

    public void AddDataGridRows(IEnumerable<object[]> rows)
    {
        var dataRows = new List<DataGridViewRow>();
        foreach (var values in rows)
        {
            var row = new DataGridViewRow();
            row.CreateCells(dataGrid, values);
            dataRows.Add(row);
        }

        if (dataRows.Count > 0)
            dataGrid.Rows.AddRange(dataRows.ToArray());
    }

    public int DataGridRowCount => dataGrid.Rows.Count;
    public int DataGridColumnCount => dataGrid.Columns.Count;

    public object GetDataGridCellValue(int row, int column)
    {
        return dataGrid.Rows[row].Cells[column].Value ?? string.Empty;
    }

    public string GetDataGridColumnHeader(int column)
    {
        return dataGrid.Columns[column].HeaderText;
    }

    public void ClearWordView()
    {
        wordView.Clear();
    }

    public void AppendToWordView(string text, Font? font = null, Color? color = null)
    {
        if (string.IsNullOrEmpty(text))
            return;

        wordView.SelectionStart = wordView.TextLength;
        if (font != null)
            wordView.SelectionFont = font;
        if (color != null)
            wordView.SelectionColor = color.Value;

        wordView.SelectedText = text;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string WordViewText
    {
        get => wordView.Text;
        set => wordView.Text = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool WordViewReadOnly
    {
        get => wordView.ReadOnly;
        set => wordView.ReadOnly = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    string IReaderHost.StatusText
    {
        get => statusLabel.Text ?? string.Empty;
        set => statusLabel.Text = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    string IReaderHost.StatsText
    {
        get => statusStats.Text ?? string.Empty;
        set
        {
            _lastRowColStats = value;
            UpdateStatsText();
        }
    }

    private void UpdateStatsText()
    {
        if (dataGrid.Visible && dataGrid.SelectedCells.Count > 0)
        {
            var cell = dataGrid.SelectedCells[0];
            var colName = cell.OwningColumn?.HeaderText ?? "?";
            statusStats.Text = $"第 {cell.RowIndex + 1} 行, {colName} 列 | {_lastRowColStats}";
        }
        else
        {
            statusStats.Text = _lastRowColStats;
        }
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    string IReaderHost.Subtitle
    {
        get => lblSubtitle.Text;
        set => lblSubtitle.Text = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    string IReaderHost.FormTitle
    {
        get => Text;
        set => Text = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    bool IReaderHost.SaveEnabled
    {
        get => btnSave.Enabled;
        set => btnSave.Enabled = value;
    }

    public bool AskLargeFile(string message)
    {
        var result = MessageBox.Show(message, "大文件警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        return result == DialogResult.Yes;
    }

    public void ShowError(string message)
    {
        MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    public void ShowInfo(string message)
    {
        MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void PumpEvents()
    {
        Application.DoEvents();
    }

    // ═══════════════════ DataGrid Styling ═══════════════════

    private void SetupDataGridStyle()
    {
        dataGrid.EnableHeadersVisualStyles = false;

        dataGrid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = ThemeConstants.Primary,
            ForeColor = ThemeConstants.TextOnPrimary,
            Font = ThemeConstants.FontBodyBold,
            Padding = new Padding(ThemeConstants.CellPadding, 0, ThemeConstants.CellPadding, 0),
            Alignment = DataGridViewContentAlignment.MiddleLeft,
            SelectionBackColor = ThemeConstants.Primary
        };
        dataGrid.ColumnHeadersHeight = ThemeConstants.ColumnHeaderHeight;
        dataGrid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

        dataGrid.DefaultCellStyle = new DataGridViewCellStyle
        {
            Font = ThemeConstants.FontBody,
            ForeColor = ThemeConstants.TextPrimary,
            Padding = new Padding(ThemeConstants.CellPadding, 0, ThemeConstants.CellPadding, 0),
            SelectionBackColor = ThemeConstants.SelectionBack,
            SelectionForeColor = ThemeConstants.SelectionFore,
            WrapMode = DataGridViewTriState.False
        };

        dataGrid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = ThemeConstants.BackgroundAltRow,
            Font = ThemeConstants.FontBody,
            ForeColor = ThemeConstants.TextPrimary,
            Padding = new Padding(ThemeConstants.CellPadding, 0, ThemeConstants.CellPadding, 0),
            SelectionBackColor = ThemeConstants.SelectionBack,
            SelectionForeColor = ThemeConstants.SelectionFore,
            WrapMode = DataGridViewTriState.False
        };

        dataGrid.RowTemplate.Height = ThemeConstants.RowHeight;
        dataGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

        dataGrid.CellFormatting += DataGrid_CellFormatting;
        dataGrid.CellToolTipTextNeeded += DataGrid_CellToolTipTextNeeded;
        dataGrid.SelectionChanged += DataGrid_SelectionChanged;
    }

    // ═══════════════════ File opening ═══════════════════

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "所有支持的文件|*.xlsx;*.xls;*.xlsm;*.docx;*.pdf;*.txt;*.csv|Excel 文件|*.xlsx;*.xls;*.xlsm|Word 文件|*.docx|PDF 文件|*.pdf|文本文件|*.txt|CSV 文件|*.csv|所有文件|*.*",
            Title = "打开文档文件"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
            OpenFile(dialog.FileName);
    }

    private void OpenFile(string path)
    {
        var ext = Path.GetExtension(path).ToLower();

        if (DocumentReaderFactory.TryGetReader(ext, out var reader))
        {
            _currentReader = reader;
            _excelReader = reader as ExcelDocumentReader;
            reader!.Load(path, this);
            if (dataGrid.Visible)
                ApplyColumnSizing();
        }
        else
        {
            ShowInfo("不支持的文件格式。");
        }
    }

    // ═══════════════════ Sheet switching ═══════════════════

    private void CmbSheet_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (cmbSheet.SelectedItem is string sheetName && _excelReader != null)
        {
            _excelReader.LoadSheetByName(sheetName, this);
            ApplyColumnSizing();
        }
    }

    private void ApplyColumnSizing()
    {
        if (dataGrid.Columns.Count == 0) return;

        dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

        // 按内容采样计算每列宽度
        foreach (DataGridViewColumn col in dataGrid.Columns)
        {
            int idealWidth = MeasureTextWidth(col.HeaderText, dataGrid.ColumnHeadersDefaultCellStyle.Font)
                + ThemeConstants.CellPadding * 2 + 24; // 24px 余量给排序箭头

            int sampleLimit = Math.Min(dataGrid.Rows.Count, 100);
            for (int r = 0; r < sampleLimit; r++)
            {
                var val = dataGrid.Rows[r].Cells[col.Index].Value?.ToString();
                if (!string.IsNullOrEmpty(val))
                {
                    int w = MeasureTextWidth(val, dataGrid.DefaultCellStyle.Font)
                        + ThemeConstants.CellPadding * 2;
                    if (w > idealWidth) idealWidth = w;
                }
            }

            col.Width = Math.Min(idealWidth, ThemeConstants.MaxColumnPixelWidth);
        }

        // 最后一列填充剩余空间
        var lastCol = dataGrid.Columns[^1];
        lastCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        lastCol.MinimumWidth = Math.Min(lastCol.Width, ThemeConstants.MaxColumnPixelWidth);
    }

    private static int MeasureTextWidth(string text, Font font)
    {
        using var g = Graphics.FromHwnd(IntPtr.Zero);
        var size = TextRenderer.MeasureText(g, text, font);
        return size.Width;
    }

    // ═══════════════════ Save ═══════════════════

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        if (_currentReader is not ISaveableDocumentReader saveable) return;

        try
        {
            saveable.Save(this);
        }
        catch (Exception ex)
        {
            Cursor = Cursors.Default;
            ShowError($"保存失败:\n{ex.Message}");
        }
    }

    // ═══════════════════ Cell display helpers ═══════════════════

    private void DataGrid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.Value is string text)
        {
            // 将换行符替换为空格，强制单行显示，防止行高被撑大
            var normalized = text.ReplaceLineEndings(" ");
            if (normalized.Length > 80)
            {
                e.Value = normalized[..80] + "…";
            }
            else
            {
                e.Value = normalized;
            }
            e.FormattingApplied = true;
        }
    }

    private void DataGrid_CellToolTipTextNeeded(object? sender, DataGridViewCellToolTipTextNeededEventArgs e)
    {
        if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
        {
            var val = dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (val is string text && !string.IsNullOrWhiteSpace(text))
            {
                e.ToolTipText = text;
            }
        }
    }

    private void DataGrid_SelectionChanged(object? sender, EventArgs e)
    {
        UpdateStatsText();
    }

    // ═══════════════════ Drag & Drop ═══════════════════

    private void Form_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files?.Length == 1 && DocumentReaderFactory.SupportedExtensions.Contains(Path.GetExtension(files[0]).ToLower()))
            {
                e.Effect = DragDropEffects.Copy;
                return;
            }
        }
        e.Effect = DragDropEffects.None;
    }

    private void Form_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            OpenFile(files[0]);
    }
}
