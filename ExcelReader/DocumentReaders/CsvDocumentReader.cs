using System.Text;

namespace ExcelReader;

/// <summary>
/// Reads CSV files into the grid view.
/// </summary>
public class CsvDocumentReader : IDocumentReader, ISaveableDocumentReader
{
    public bool CanSave => true;

    private string? _currentPath;
    private Encoding _encoding = new UTF8Encoding(false);
    private char _delimiter = ',';

    public void Load(string path, IReaderHost host)
    {
        try
        {
            host.ShowExcelView();
            host.SaveEnabled = true;
            host.SetSheets(Array.Empty<string>());
            host.BeginUpdate();
            try
            {
                host.ClearDataGrid();
                host.Cursor = Cursors.WaitCursor;
                host.StatusText = "正在加载 CSV...";
                host.PumpEvents();

                _currentPath = path;
                _encoding = TextFileHelper.DetectEncoding(path);
                var text = File.ReadAllText(path, _encoding);
                _delimiter = DetectDelimiter(text);
                var rows = ParseDelimitedText(text, _delimiter);

                if (rows.Count == 0)
                {
                    host.StatsText = $"空文件 | {GetDelimiterName(_delimiter)} | {TextFileHelper.GetEncodingDisplayName(_encoding)}";
                }
                else
                {
                    var headers = rows[0];
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var header = string.IsNullOrWhiteSpace(headers[i]) ? $"列{i + 1}" : headers[i];
                        host.AddDataGridColumn($"col{i + 1}", header);
                    }

                    foreach (var row in rows.Skip(1))
                    {
                        var values = new object[headers.Count];
                        for (int i = 0; i < headers.Count; i++)
                        {
                            values[i] = i < row.Count ? row[i] : string.Empty;
                        }
                        host.AddDataGridRow(values);
                    }

                    host.StatsText = $"{Math.Max(rows.Count - 1, 0)} 行 × {headers.Count} 列 | {GetDelimiterName(_delimiter)} | {TextFileHelper.GetEncodingDisplayName(_encoding)}";
                }
            }
            finally
            {
                host.EndUpdate();
                host.Cursor = Cursors.Default;
            }

            var fileName = Path.GetFileName(path);
            host.Subtitle = fileName;
            host.StatusText = $"已打开: {fileName}";
            host.FormTitle = $"Office 阅读器 - {fileName}";
        }
        catch (Exception ex)
        {
            host.Cursor = Cursors.Default;
            host.ShowError($"无法打开 CSV 文件:\n{ex.Message}");
        }
    }

    public void Save(IReaderHost host)
    {
        if (string.IsNullOrEmpty(_currentPath)) return;

        try
        {
            host.Cursor = Cursors.WaitCursor;
            host.StatusText = "正在保存 CSV...";
            host.PumpEvents();

            var rows = new List<string>();
            var headers = new List<string>();
            for (int c = 0; c < host.DataGridColumnCount; c++)
            {
                headers.Add(EscapeField(host.GetDataGridColumnHeader(c), _delimiter));
            }
            rows.Add(string.Join(_delimiter, headers));

            for (int r = 0; r < host.DataGridRowCount; r++)
            {
                var values = new List<string>();
                for (int c = 0; c < host.DataGridColumnCount; c++)
                {
                    var value = host.GetDataGridCellValue(r, c)?.ToString() ?? string.Empty;
                    values.Add(EscapeField(value, _delimiter));
                }
                rows.Add(string.Join(_delimiter, values));
            }

            File.WriteAllText(_currentPath, string.Join(Environment.NewLine, rows), _encoding);
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

    private static List<List<string>> ParseDelimitedText(string text, char delimiter)
    {
        var rows = new List<List<string>>();
        var currentRow = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (int i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (ch == delimiter && !inQuotes)
            {
                currentRow.Add(current.ToString());
                current.Clear();
            }
            else if ((ch == '\n' || ch == '\r') && !inQuotes)
            {
                if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;

                currentRow.Add(current.ToString());
                current.Clear();
                rows.Add(currentRow);
                currentRow = new List<string>();
            }
            else
            {
                current.Append(ch);
            }
        }

        if (current.Length > 0 || currentRow.Count > 0)
        {
            currentRow.Add(current.ToString());
            rows.Add(currentRow);
        }

        return rows;
    }

    private static char DetectDelimiter(string text)
    {
        var candidates = new[] { ',', ';', '\t', '|' };
        var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
            .Take(10)
            .ToArray();

        if (lines.Length == 0)
            return ',';

        char best = ',';
        int bestScore = int.MinValue;

        foreach (var candidate in candidates)
        {
            var counts = lines.Select(line => CountDelimiter(line, candidate)).ToArray();
            int max = counts.Max();
            int min = counts.Min();
            int score = counts.Sum() - (max - min);
            if (max == 0)
                continue;

            if (score > bestScore)
            {
                bestScore = score;
                best = candidate;
            }
        }

        return best;
    }

    private static int CountDelimiter(string line, char delimiter)
    {
        int count = 0;
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (line[i] == delimiter && !inQuotes)
            {
                count++;
            }
        }
        return count;
    }

    private static string EscapeField(string value, char delimiter)
    {
        if (value.Contains('"'))
            value = value.Replace("\"", "\"\"");

        if (value.Contains(delimiter) || value.Contains('\n') || value.Contains('\r') || value.Contains('"'))
            return $"\"{value}\"";

        return value;
    }

    private static string GetDelimiterName(char delimiter)
    {
        return delimiter switch
        {
            ',' => "逗号分隔",
            ';' => "分号分隔",
            '\t' => "Tab 分隔",
            '|' => "竖线分隔",
            _ => $"分隔符 {delimiter}"
        };
    }
}
