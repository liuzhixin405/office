using System.Collections.Frozen;

namespace ExcelReader;

/// <summary>
/// Factory that maps file extensions to appropriate IDocumentReader instances.
/// New formats can be added by extending the dictionary without modifying MainForm.
/// </summary>
public static class DocumentReaderFactory
{
    private static readonly FrozenDictionary<string, IDocumentReader> Readers =
        new Dictionary<string, IDocumentReader>(StringComparer.OrdinalIgnoreCase)
        {
            [".xlsx"] = new ExcelDocumentReader(),
            [".xls"] = new ExcelDocumentReader(),
            [".xlsm"] = new ExcelDocumentReader(),
            [".docx"] = new WordDocumentReader(),
            [".pdf"] = new PdfDocumentReader(),
            [".txt"] = new TextDocumentReader(),
            [".csv"] = new CsvDocumentReader()
        }.ToFrozenDictionary();

    public static readonly FrozenSet<string> SupportedExtensions = Readers.Keys.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Try to get a reader for the given file extension (case-insensitive).</summary>
    public static bool TryGetReader(string extension, out IDocumentReader? reader)
    {
        return Readers.TryGetValue(extension, out reader);
    }
}
