namespace ExcelReader;

/// <summary>
/// Abstraction of the UI host that document readers interact with.
/// This decouples readers from the concrete Form implementation.
/// </summary>
public interface IReaderHost
{
    // ── UI visibility control ──────────────────────────
    void ShowExcelView();
    void ShowWordPdfView();

    // ── Update batching (reduces flicker during bulk loading) ──
    void BeginUpdate();
    void EndUpdate();

    // ── Sheet combo ────────────────────────────────────
    void SetSheets(IReadOnlyList<string> sheetNames);
    string? SelectedSheet { get; }

    // ── Data grid ──────────────────────────────────────
    void ClearDataGrid();
    void AddDataGridColumn(string name, string headerText);
    void AddDataGridRow(object[] values);
    void AddDataGridRows(IEnumerable<object[]> rows);
    int DataGridRowCount { get; }
    int DataGridColumnCount { get; }
    object GetDataGridCellValue(int row, int column);
    string GetDataGridColumnHeader(int column);

    // ── Rich text view ─────────────────────────────────
    void ClearWordView();
    void AppendToWordView(string text, Font? font = null, Color? color = null);
    string WordViewText { get; set; }
    bool WordViewReadOnly { get; set; }

    // ── Cursor ─────────────────────────────────────────
    Cursor Cursor { get; set; }

    // ── Status ─────────────────────────────────────────
    string StatusText { get; set; }
    string StatsText { get; set; }
    string Subtitle { get; set; }
    string FormTitle { get; set; }

    // ── Save ───────────────────────────────────────────
    bool SaveEnabled { get; set; }

    // ── Dialogs ────────────────────────────────────────
    bool AskLargeFile(string message);
    void ShowError(string message);
    void ShowInfo(string message);
    void PumpEvents();
}
