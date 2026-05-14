using System.Text;

namespace ExcelReader;

/// <summary>
/// Reads plain text (.txt) files into the text view.
/// </summary>
public class TextDocumentReader : IDocumentReader, ISaveableDocumentReader
{
    public bool CanSave => true;

    private string? _currentPath;
    private Encoding _encoding = new UTF8Encoding(false);

    public void Load(string path, IReaderHost host)
    {
        try
        {
            host.ShowWordPdfView();
            host.WordViewReadOnly = false;
            host.SaveEnabled = true;
            host.ClearWordView();
            host.Cursor = Cursors.WaitCursor;
            host.StatusText = "正在加载文本文件...";
            host.PumpEvents();

            _currentPath = path;
            _encoding = TextFileHelper.DetectEncoding(path);
            var text = File.ReadAllText(path, _encoding);

            host.BeginUpdate();
            try
            {
                host.WordViewText = text;
                var lineCount = text.Length == 0 ? 0 : text.Count(c => c == '\n') + 1;
                host.StatsText = $"{lineCount} 行 | {TextFileHelper.GetEncodingDisplayName(_encoding)}";
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
            host.ShowError($"无法打开文本文件:\n{ex.Message}");
        }
    }

    public void Save(IReaderHost host)
    {
        if (string.IsNullOrEmpty(_currentPath)) return;

        try
        {
            host.Cursor = Cursors.WaitCursor;
            host.StatusText = "正在保存文本文件...";
            host.PumpEvents();
            File.WriteAllText(_currentPath, host.WordViewText, _encoding);
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
