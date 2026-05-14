using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelReader;

/// <summary>
/// Reads Word (.docx) files via OpenXml.
/// </summary>
public class WordDocumentReader : IDocumentReader
{
    public bool CanSave => false;

    public void Load(string path, IReaderHost host)
    {
        try
        {
            host.ShowWordPdfView();
            host.SaveEnabled = false;
            host.ClearWordView();

            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                host.AppendToWordView("(空文档)");
                return;
            }

            host.Cursor = Cursors.WaitCursor;
            host.StatusText = "正在加载...";
            host.PumpEvents();

            host.BeginUpdate();
            try
            {
                int paragraphCount = 0;
                foreach (var element in body.ChildElements)
                {
                    if (element is W.Paragraph para)
                    {
                        AppendParagraph(para, host);
                        paragraphCount++;
                    }
                    else if (element is W.Table table)
                    {
                        AppendTable(table, host);
                    }
                }

                host.StatsText = $"{paragraphCount} 个段落";
            }
            finally
            {
                host.EndUpdate();
            }

            var fileName = Path.GetFileName(path);
            host.Subtitle = fileName;
            host.StatusText = $"已打开: {fileName}";
            host.FormTitle = $"Office 阅读器 - {fileName}";
        }
        catch (Exception ex)
        {
            host.Cursor = Cursors.Default;
            host.ShowError($"无法打开文件:\n{ex.Message}");
        }
    }

    private static void AppendParagraph(W.Paragraph para, IReaderHost host)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        bool isHeading = styleId?.StartsWith("Heading") == true;
        int headingLevel = isHeading && styleId!.Length > 7 ? (styleId[7] - '0') : 0;

        if (isHeading && headingLevel is >= 1 and <= 3)
        {
            Font font = headingLevel switch
            {
                1 => ThemeConstants.FontHeading1,
                2 => ThemeConstants.FontHeading2,
                _ => ThemeConstants.FontHeading3
            };
            Color color = ThemeConstants.Primary;

            host.AppendToWordView(string.Empty, font, color); // set format
            AppendRuns(para, host);
        }
        else
        {
            host.AppendToWordView(string.Empty, ThemeConstants.FontParagraph, ThemeConstants.TextPrimary); // set format

            bool appended = false;
            foreach (var run in para.Elements<W.Run>())
            {
                bool isBold = run.RunProperties?.Bold?.Val?.Value == true;
                if (isBold)
                {
                    host.AppendToWordView(run.InnerText, ThemeConstants.FontParagraphBold);
                }
                else
                {
                    host.AppendToWordView(run.InnerText);
                }
                appended = true;
            }

            if (!appended)
            {
                host.AppendToWordView(para.InnerText);
            }
        }

        host.AppendToWordView("\n");
    }

    private static void AppendRuns(W.Paragraph para, IReaderHost host)
    {
        foreach (var run in para.Elements<W.Run>())
            host.AppendToWordView(run.InnerText);
    }

    private static void AppendTable(W.Table table, IReaderHost host)
    {
        host.AppendToWordView(string.Empty, ThemeConstants.FontTableMono, ThemeConstants.TextTableBorder); // set format

        foreach (var row in table.Elements<W.TableRow>())
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements<W.TableCell>())
                cells.Add(cell.InnerText.Trim().Replace("\n", " "));

            host.AppendToWordView("  " + string.Join("  │  ", cells) + "\n");
        }

        host.AppendToWordView("\n");
    }
}
