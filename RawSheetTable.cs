namespace CofyDev.Xml.Doc;

public sealed class RawSheetTable : ISheetTable
{
    public string Name { get; }
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    public RawSheetTable(string name, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        Name = name ?? string.Empty;
        Rows = rows ?? System.Array.Empty<IReadOnlyList<string>>();
    }
}
