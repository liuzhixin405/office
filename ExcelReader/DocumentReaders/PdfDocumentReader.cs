using System.Text;
using UglyToad.PdfPig;

namespace ExcelReader;

/// <summary>
/// Reads PDF files via PdfPig.
/// </summary>
public class PdfDocumentReader : IDocumentReader
{
    public bool CanSave => false;

    public void Load(string path, IReaderHost host)
    {
        try
        {
            host.ShowWordPdfView();
            host.SaveEnabled = false;
            host.ClearWordView();

            host.Cursor = Cursors.WaitCursor;
            host.StatusText = "正在解析 PDF...";
            host.PumpEvents();

            using var document = PdfDocument.Open(path);
            int pageCount = document.NumberOfPages;

            host.BeginUpdate();
            try
            {
                foreach (var page in document.GetPages())
                {
                    // Page number separator
                    host.AppendToWordView(
                        $"\n--- 第 {page.Number} 页 ---\n\n",
                        ThemeConstants.FontPageNumber,
                        ThemeConstants.TextMuted);

                    // Body text
                    host.AppendToWordView(page.Text, ThemeConstants.FontParagraph, ThemeConstants.TextPrimary);
                    host.AppendToWordView("\n");
                }

                host.StatsText = $"共 {pageCount} 页";
            }
            finally
            {
                host.EndUpdate();
            }

            var fileName = Path.GetFileName(path);
            host.Subtitle = fileName;
            host.StatusText = $"已打开: {fileName}";
            host.FormTitle = $"Office/PDF 阅读器 - {fileName}";
        }
        catch (Exception ex)
        {
            host.Cursor = Cursors.Default;
            host.ShowError($"无法打开 PDF 文件:\n{ex.Message}");
        }
    }
}
