namespace CellShell.Core;

/// <summary>
/// Pure logic for the spreadsheet grid — no WPF dependencies, fully unit-testable.
/// </summary>
public class SpreadsheetModel
{
    public const int RowNumberWidth = 40;
    public const int MinVisibleRows = 70;
    public const int MinColWidth = 50;
    public const int MinRowHeight = 16;
    public const double DefaultColWidth = 80;

    public List<CellData> Rows { get; } = new();
    public int ActiveRowIndex { get; private set; }
    public List<double> ColWidths { get; } = new() { 200, 300, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80 };
    public int DefaultRowHeight { get; private set; } = 22;
    public double CellFontSize { get; private set; } = 11;

    private readonly Dictionary<int, int> _rowHeights = new();

    /// <summary>Gets the column letter(s) for a given index: 0=A, 25=Z, 26=AA, etc.</summary>
    public static string GetColLetter(int index)
    {
        var result = "";
        index++;
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    /// <summary>Ensures enough columns exist so TotalWidth >= viewportWidth.</summary>
    public void EnsureColumnsForWidth(double viewportWidth)
    {
        while (TotalWidth < viewportWidth)
            ColWidths.Add(DefaultColWidth);
    }

    public int ColCount => ColWidths.Count;

    public double TotalWidth
    {
        get
        {
            double sum = 0;
            foreach (var w in ColWidths) sum += w;
            return sum;
        }
    }

    public int GetRowHeight(int row) =>
        _rowHeights.TryGetValue(row, out var h) ? h : DefaultRowHeight;

    public void SetRowHeight(int row, int height) =>
        _rowHeights[row] = Math.Max(MinRowHeight, height);

    // ─── Row selection ───────────────────────────────────────────

    /// <summary>Select any row (existing, new, or empty filler).</summary>
    public void SelectRow(int rowIndex)
    {
        if (rowIndex < 0) rowIndex = 0;
        ActiveRowIndex = rowIndex;
    }

    /// <summary>Move active to the first empty row after all data.</summary>
    public void MoveToNextEmptyRow()
    {
        ActiveRowIndex = Rows.Count;
    }

    // ─── Command submission ──────────────────────────────────────

    /// <summary>
    /// Submit a command at the current active row.
    /// Returns the CellData to be executed.
    /// Pads with empty rows if there's a gap.
    /// </summary>
    public CellData SubmitCommand(string command)
    {
        if (ActiveRowIndex < Rows.Count)
        {
            // Re-running existing row
            var data = Rows[ActiveRowIndex];
            data.Command = command;
            data.Status = CellStatus.Running;
            data.Output = "";
            return data;
        }

        // Pad with empty rows if clicking beyond current data
        while (Rows.Count < ActiveRowIndex)
        {
            Rows.Add(new CellData
            {
                RowNumber = Rows.Count + 1,
                Status = CellStatus.Empty
            });
        }

        var newData = new CellData
        {
            RowNumber = ActiveRowIndex + 1,
            Command = command,
            Status = CellStatus.Running
        };
        Rows.Add(newData);
        return newData;
    }

    // ─── Font size ───────────────────────────────────────────────

    public void SetFontSize(double size)
    {
        CellFontSize = Math.Clamp(size, 8, 28);
        DefaultRowHeight = Math.Max(MinRowHeight, (int)(CellFontSize * 2));
    }

    // ─── Output formatting ───────────────────────────────────────

    public static string FormatOutput(string output, CellStatus status)
    {
        if (status == CellStatus.Running && string.IsNullOrEmpty(output))
            return "Calculating...";
        if (string.IsNullOrEmpty(output)) return string.Empty;

        string formatted;
        if (output.Contains('\n'))
        {
            var lines = output.Split('\n');
            var nonEmpty = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
            if (nonEmpty.Length > 1)
                formatted = nonEmpty[0].TrimEnd('\r') + $" [{nonEmpty.Length} lines]";
            else if (nonEmpty.Length == 1)
                formatted = nonEmpty[0].TrimEnd('\r');
            else
                formatted = string.Empty;
        }
        else
        {
            formatted = output;
        }

        if (status == CellStatus.Running && !string.IsNullOrEmpty(formatted))
            return formatted + " ...";

        return formatted;
    }
}
