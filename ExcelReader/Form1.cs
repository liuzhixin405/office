using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelReader;

public partial class Form1 : Form
{
    private XLWorkbook? _workbook;
    public Form1()
    {
        InitializeComponent();
        SetupDataGridStyle();
    }

    private void SetupDataGridStyle()
    {
        dataGrid.EnableHeadersVisualStyles = false;
        dataGrid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(43, 87, 154),
            ForeColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
            Padding = new Padding(8, 0, 8, 0),
            Alignment = DataGridViewContentAlignment.MiddleLeft
        };
        dataGrid.ColumnHeadersHeight = 36;
        dataGrid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        dataGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(43, 87, 154);

        dataGrid.DefaultCellStyle = new DataGridViewCellStyle
        {
            Font = new Font("Microsoft YaHei UI", 9f),
            ForeColor = Color.FromArgb(50, 50, 50),
            Padding = new Padding(8, 0, 8, 0),
            SelectionBackColor = Color.FromArgb(220, 235, 252),
            SelectionForeColor = Color.FromArgb(50, 50, 50)
        };
        dataGrid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(248, 249, 250),
            Font = new Font("Microsoft YaHei UI", 9f),
            ForeColor = Color.FromArgb(50, 50, 50),
            Padding = new Padding(8, 0, 8, 0),
            SelectionBackColor = Color.FromArgb(220, 235, 252),
            SelectionForeColor = Color.FromArgb(50, 50, 50)
        };

        dataGrid.RowTemplate.Height = 28;
    }

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Office 文件|*.xlsx;*.xls;*.xlsm;*.docx|Excel 文件|*.xlsx;*.xls;*.xlsm|Word 文件|*.docx|所有文件|*.*",
            Title = "打开 Office 文件"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            OpenFile(dialog.FileName);
        }
    }

    private void OpenFile(string path)
    {
        var ext = Path.GetExtension(path).ToLower();
        if (ext is ".xlsx" or ".xls" or ".xlsm")
        {
            LoadExcel(path);
        }
        else if (ext == ".docx")
        {
            LoadWord(path);
        }
        else
        {
            MessageBox.Show("不支持的文件格式。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // ═══════════════════ Excel ═══════════════════

    private void LoadExcel(string path)
    {
        try
        {
            _workbook?.Dispose();
            _workbook = new XLWorkbook(path);

            cmbSheet.Items.Clear();
            foreach (var sheet in _workbook.Worksheets)
            {
                cmbSheet.Items.Add(sheet.Name);
            }

            cmbSheet.Enabled = cmbSheet.Items.Count > 0;
            if (cmbSheet.Items.Count > 0)
            {
                cmbSheet.SelectedIndex = 0;
            }

            wordView.Visible = false;
            dataGrid.Visible = true;
            lblSheetLabel.Visible = true;
            cmbSheet.Visible = true;

            var name = Path.GetFileName(path);
            lblSubtitle.Text = name;
            statusLabel.Text = $"已打开: {name}";
            Text = $"Office 阅读器 - {name}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"无法打开文件:\n{ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void CmbSheet_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (cmbSheet.SelectedItem is string sheetName)
        {
            LoadSheet(sheetName);
        }
    }

    private void LoadSheet(string sheetName)
    {
        if (_workbook == null) return;

        var sheet = _workbook.Worksheet(sheetName);
        var range = sheet.RangeUsed();
        dataGrid.Columns.Clear();

        if (range == null)
        {
            statusStats.Text = "空表";
            return;
        }

        var firstRow = range.FirstRow().RowNumber();
        var lastRow = range.LastRow().RowNumber();
        var firstCol = range.FirstColumn().ColumnNumber();
        var lastCol = range.LastColumn().ColumnNumber();

        var headerRow = sheet.Row(firstRow);
        for (int c = firstCol; c <= lastCol; c++)
        {
            dataGrid.Columns.Add($"col{c}", headerRow.Cell(c).GetString());
        }

        dataGrid.Rows.Clear();
        var rows = lastRow - firstRow;
        if (rows > 100_000)
        {
            var result = MessageBox.Show(
                $"该工作表有 {rows:N0} 行数据，加载可能需要较长时间。是否继续？",
                "大文件警告",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.No) return;
        }

        Cursor = Cursors.WaitCursor;
        statusLabel.Text = $"正在加载 \"{sheetName}\"...";
        Application.DoEvents();

        for (int r = firstRow + 1; r <= lastRow; r++)
        {
            var row = sheet.Row(r);
            var values = new object[lastCol - firstCol + 1];
            for (int c = firstCol; c <= lastCol; c++)
            {
                values[c - firstCol] = row.Cell(c).GetString();
            }
            dataGrid.Rows.Add(values);
        }

        Cursor = Cursors.Default;
        statusLabel.Text = $"工作表: {sheetName}";
        statusStats.Text = $"{lastRow - firstRow} 行 × {lastCol - firstCol + 1} 列";

        foreach (DataGridViewColumn col in dataGrid.Columns)
        {
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            var width = col.Width;
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            col.Width = Math.Min(width, 400);
        }
    }

    // ═══════════════════ Word ═══════════════════

    private void LoadWord(string path)
    {
        try
        {
            _workbook?.Dispose();
            _workbook = null;

            dataGrid.Visible = false;
            wordView.Visible = true;
            lblSheetLabel.Visible = false;
            cmbSheet.Visible = false;

            wordView.Clear();

            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document.Body;
            if (body == null)
            {
                wordView.Text = "(空文档)";
                return;
            }

            Cursor = Cursors.WaitCursor;
            statusLabel.Text = "正在加载...";
            Application.DoEvents();

            var paragraphCount = 0;
            foreach (var element in body.ChildElements)
            {
                if (element is W.Paragraph para)
                {
                    AppendParagraph(para);
                    paragraphCount++;
                }
                else if (element is W.Table table)
                {
                    AppendTable(table);
                }
            }

            Cursor = Cursors.Default;

            var name = Path.GetFileName(path);
            lblSubtitle.Text = name;
            statusLabel.Text = $"已打开: {name}";
            statusStats.Text = $"{paragraphCount} 个段落";
            Text = $"Office 阅读器 - {name}";

            // Scroll to top
            wordView.SelectionStart = 0;
            wordView.ScrollToCaret();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"无法打开文件:\n{ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AppendParagraph(W.Paragraph para)
    {
        // Check for headings
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var isHeading = styleId?.StartsWith("Heading") == true;
        var headingLevel = isHeading && styleId!.Length > 7
            ? (styleId[7] - '0')
            : 0;

        wordView.SelectionStart = wordView.TextLength;

        if (isHeading && headingLevel >= 1 && headingLevel <= 3)
        {
            float fontSize = headingLevel switch
            {
                1 => 18f,
                2 => 14f,
                _ => 12f
            };
            wordView.SelectionFont = new Font("Microsoft YaHei UI", fontSize, FontStyle.Bold);
            wordView.SelectionColor = Color.FromArgb(43, 87, 154);
            AppendRuns(para);
        }
        else
        {
            wordView.SelectionFont = new Font("Microsoft YaHei UI", 10.5f);
            wordView.SelectionColor = Color.FromArgb(50, 50, 50);

            var hasBold = false;
            foreach (var run in para.Elements<W.Run>())
            {
                var isBold = run.RunProperties?.Bold?.Val?.Value == true;
                if (isBold)
                {
                    wordView.SelectionFont = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold);
                    wordView.SelectedText = run.InnerText;
                    wordView.SelectionFont = new Font("Microsoft YaHei UI", 10.5f);
                    hasBold = true;
                }
                else
                {
                    wordView.SelectedText = run.InnerText;
                    hasBold = true;
                }
            }

            if (!hasBold)
            {
                wordView.SelectedText = para.InnerText;
            }
        }

        wordView.SelectedText = "\n";
    }

    private void AppendRuns(W.Paragraph para)
    {
        foreach (var run in para.Elements<W.Run>())
        {
            wordView.SelectedText = run.InnerText;
        }
    }

    private void AppendTable(W.Table table)
    {
        wordView.SelectionStart = wordView.TextLength;
        wordView.SelectionFont = new Font("Consolas", 9f);
        wordView.SelectionColor = Color.FromArgb(80, 80, 80);

        foreach (var row in table.Elements<W.TableRow>())
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements<W.TableCell>())
            {
                cells.Add(cell.InnerText.Trim().Replace("\n", " "));
            }
            wordView.SelectedText = "  " + string.Join("  │  ", cells) + "\n";
        }

        wordView.SelectedText = "\n";
    }

    // ═══════════════════ Drag & Drop ═══════════════════

    private static readonly HashSet<string> SupportedExtensions = new()
    {
        ".xlsx", ".xls", ".xlsm", ".docx"
    };

    private void Form_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files?.Length == 1 && SupportedExtensions.Contains(Path.GetExtension(files[0]).ToLower()))
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
        {
            OpenFile(files[0]);
        }
    }
}
