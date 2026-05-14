namespace ExcelReader;

/// <summary>
/// Represents a document reader that can load a specific file format.
/// </summary>
public interface IDocumentReader
{
    /// <summary>Whether this reader supports saving edits back to the file.</summary>
    bool CanSave { get; }

    /// <summary>Load the document and populate the host UI.</summary>
    void Load(string path, IReaderHost host);
}

/// <summary>
/// Optional contract for readers that can persist edits.
/// </summary>
public interface ISaveableDocumentReader
{
    void Save(IReaderHost host);
}