using System;
using System.Collections.Generic;

namespace CofyDev.Xml.Doc;

public static class SheetTableConverter
{
    public static DataTable ToDataTable(ISheetTable table)
    {
        if (table == null)
            throw new ArgumentNullException(nameof(table));

        var rows = table.Rows;
        if (rows.Count == 0)
            return new DataTable();

        var headerRow = rows[0];
        if (headerRow == null || headerRow.Count == 0)
            return new DataTable();

        var headerMappings = ParseHeaderMappings(headerRow);
        if (headerMappings.Count == 0)
            return new DataTable();

        var data = new DataTable();
        for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            if (row == null || IsEmptyRow(row))
                continue;

            var dataRow = new DataRow(headerMappings.Count);
            foreach (var mapping in headerMappings)
            {
                if (mapping.IsDuplicate)
                {
                    var values = CollectDuplicateValues(row, mapping.Indices);
                    if (values.Length > 0)
                        dataRow[mapping.Name] = values;
                }
                else
                {
                    var value = GetValue(row, mapping.SingleIndex);
                    if (value != null)
                        dataRow[mapping.Name] = value;
                }
            }

            data.Add(dataRow);
        }

        return data;
    }

    private static IReadOnlyList<HeaderMapping> ParseHeaderMappings(IReadOnlyList<string> headerRow)
    {
        var headerCount = headerRow.Count;
        var headerToIndices = new Dictionary<string, List<int>>(headerCount, StringComparer.Ordinal);

        for (var columnIndex = 0; columnIndex < headerCount; columnIndex++)
        {
            var header = headerRow[columnIndex]?.Trim();
            if (string.IsNullOrWhiteSpace(header))
                continue;

            if (!headerToIndices.TryGetValue(header, out var indices))
            {
                indices = new List<int>(2);
                headerToIndices[header] = indices;
            }

            indices.Add(columnIndex);
        }

        var mappings = new List<HeaderMapping>(headerToIndices.Count);
        foreach (var (header, indices) in headerToIndices)
        {
            if (indices.Count == 1)
                mappings.Add(new HeaderMapping(header, indices[0]));
            else
                mappings.Add(new HeaderMapping(header, indices.ToArray()));
        }

        return mappings;
    }

    private static bool IsEmptyRow(IReadOnlyList<string> row)
    {
        for (var i = 0; i < row.Count; i++)
        {
            if (!string.IsNullOrWhiteSpace(row[i]))
                return false;
        }

        return true;
    }

    private static string[] CollectDuplicateValues(IReadOnlyList<string> row, int[] indices)
    {
        var count = 0;
        foreach (var index in indices)
        {
            if (index < row.Count && !string.IsNullOrEmpty(row[index]))
                count++;
        }

        if (count == 0)
            return Array.Empty<string>();

        var values = new string[count];
        var valueIndex = 0;
        foreach (var index in indices)
        {
            if (index < row.Count)
            {
                var value = row[index];
                if (!string.IsNullOrEmpty(value))
                    values[valueIndex++] = value;
            }
        }

        return values;
    }

    private static string GetValue(IReadOnlyList<string> row, int index)
    {
        if (index < 0 || index >= row.Count)
            return null;

        var value = row[index];
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private readonly struct HeaderMapping
    {
        public string Name { get; }
        public bool IsDuplicate { get; }
        public int SingleIndex { get; }
        public int[] Indices { get; }

        public HeaderMapping(string name, int singleIndex)
        {
            Name = name;
            IsDuplicate = false;
            SingleIndex = singleIndex;
            Indices = null;
        }

        public HeaderMapping(string name, int[] indices)
        {
            Name = name;
            IsDuplicate = true;
            SingleIndex = -1;
            Indices = indices;
        }
    }
}
