namespace CofyDev.Xml.Doc;

public interface ISheetTable
{
    string Name { get; }
    IReadOnlyList<IReadOnlyList<string>> Rows { get; }
}
