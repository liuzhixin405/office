using ClosedXML.Excel;

namespace ExcelReader;

/// <summary>
/// Reads and optionally saves Excel (.xlsx/.xls/.xlsm) files via ClosedXML.
/// </summary>
public class ExcelDocumentReader : IDocumentReader, ISaveableDocumentReader
{
    public bool CanSave => true;

    private XLWorkbook? _workbook;
    private string? _currentPath;

    public void Load(string path, IReaderHost host)
    {
        try
        {
            _workbook?.Dispose();
            _workbook = new XLWorkbook(path);
            _currentPath = path;

            var sheetNames = _workbook.Worksheets.Select(s => s.Name).ToList();
            host.SetSheets(sheetNames);
            host.ShowExcelView();
            host.SaveEnabled = true;

            host.Subtitle = Path.GetFileName(path);
            host.StatusText = $"已打开: {Path.GetFileName(path)}";
            host.FormTitle = $"Office 阅读器 - {Path.GetFileName(path)}";

            if (sheetNames.Count > 0)
                LoadSheet(sheetNames[0], host);
        }
        catch (Exception ex)
        {
            _workbook?.Dispose();
            _workbook = null;
            _currentPath = null;
            host.SaveEnabled = false;
            host.Cursor = Cursors.Default;
            host.ShowError($"无法打开文件:\n{ex.Message}");
        }
    }

    private void LoadSheet(string sheetName, IReaderHost host)
    {
        if (_workbook == null) return;

        var sheet = _workbook.Worksheet(sheetName);
        var range = sheet.RangeUsed();

        host.BeginUpdate();
        try
        {
            host.ClearDataGrid();

            if (range == null)
            {
                host.StatsText = "空表";
                return;
            }

            var firstRow = range.FirstRow().RowNumber();
            var lastRow = range.LastRow().RowNumber();
            var firstCol = range.FirstColumn().ColumnNumber();
            var lastCol = range.LastColumn().ColumnNumber();

            // Build headers
            var headerRow = sheet.Row(firstRow);
            for (int c = firstCol; c <= lastCol; c++)
            {
                host.AddDataGridColumn($"col{c}", headerRow.Cell(c).GetString());
            }

            var rows = lastRow - firstRow;
            if (rows > ThemeConstants.LargeSheetRowThreshold)
            {
                var result = host.AskLargeFile($"该工作表有 {rows:N0} 行数据，加载可能需要较长时间。是否继续？");
                if (!result) return;
            }

            host.Cursor = Cursors.WaitCursor;
            host.StatusText = $"正在加载 \"{sheetName}\"...";
            host.PumpEvents();

            var dataRows = range.Rows(row => row.RowNumber() > firstRow);
            var batch = new List<object[]>();
            int colCount = lastCol - firstCol + 1;

            foreach (var row in dataRows)
            {
                var values = new object[colCount];
                for (int c = firstCol; c <= lastCol; c++)
                {
                    values[c - firstCol] = row.Cell(c).GetString();
                }
                batch.Add(values);
            }

            host.AddDataGridRows(batch);

            host.StatusText = $"工作表: {sheetName}";
            host.StatsText = $"{rows} 行 × {colCount} 列";
        }
        finally
        {
            host.Cursor = Cursors.Default;
            host.EndUpdate();
        }
    }

    public void LoadSheetByName(string sheetName, IReaderHost host)
    {
        LoadSheet(sheetName, host);
    }

    public void Save(IReaderHost host)
    {
        if (_workbook == null || string.IsNullOrEmpty(_currentPath)) return;

        var sheetName = host.SelectedSheet;
        if (string.IsNullOrEmpty(sheetName)) return;

        host.Cursor = Cursors.WaitCursor;
        host.StatusText = "正在保存更改...";
        host.PumpEvents();

        try
        {
            var sheet = _workbook.Worksheet(sheetName);
            int rowCount = host.DataGridRowCount;
            int colCount = host.DataGridColumnCount;

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    var val = host.GetDataGridCellValue(r, c)?.ToString() ?? "";
                    sheet.Cell(r + 2, c + 1).Value = val;
                }
            }

            _workbook.Save();
            host.StatusText = "保存完成";
            host.ShowInfo("更改已保存到原文件。");
        }
        catch (Exception ex)
        {
            host.ShowError($"保存失败:\n{ex.Message}");
        }
        finally
        {
            host.Cursor = Cursors.Default;
        }
    }
}
