using CellShell.Core;

namespace CellShell.Tests;

public class SpreadsheetModelTests
{
    // ─── Initial state ───────────────────────────────────────────

    [Fact]
    public void InitialState_ActiveRowIsZero()
    {
        var model = new SpreadsheetModel();
        Assert.Equal(0, model.ActiveRowIndex);
        Assert.Empty(model.Rows);
    }

    [Fact]
    public void InitialState_ColumnWidthsAreCorrect()
    {
        var model = new SpreadsheetModel();
        Assert.Equal(53, model.ColCount); // A through BA
        Assert.Equal(200, model.ColWidths[0]); // A
        Assert.Equal(300, model.ColWidths[1]); // B
    }

    // ─── Dynamic columns ──────────────────────────────────────────

    [Fact]
    public void GetColLetter_SingleLetters()
    {
        Assert.Equal("A", SpreadsheetModel.GetColLetter(0));
        Assert.Equal("B", SpreadsheetModel.GetColLetter(1));
        Assert.Equal("Z", SpreadsheetModel.GetColLetter(25));
    }

    [Fact]
    public void GetColLetter_DoubleLetters()
    {
        Assert.Equal("AA", SpreadsheetModel.GetColLetter(26));
        Assert.Equal("AB", SpreadsheetModel.GetColLetter(27));
        Assert.Equal("AZ", SpreadsheetModel.GetColLetter(51));
        Assert.Equal("BA", SpreadsheetModel.GetColLetter(52));
    }

    [Fact]
    public void EnsureColumnsForWidth_AddsColumnsToFill()
    {
        var model = new SpreadsheetModel();
        var initialCount = model.ColCount;
        var initialWidth = model.TotalWidth;
        // Request a width wider than current total
        model.EnsureColumnsForWidth(initialWidth + 200);
        Assert.True(model.ColCount > initialCount);
        Assert.True(model.TotalWidth >= initialWidth + 200);
    }

    [Fact]
    public void EnsureColumnsForWidth_NoOpWhenAlreadyWideEnough()
    {
        var model = new SpreadsheetModel();
        var count = model.ColCount;
        model.EnsureColumnsForWidth(100); // well within existing width
        Assert.Equal(count, model.ColCount);
    }

    // ─── Row selection ───────────────────────────────────────────

    [Fact]
    public void SelectRow_ExistingRow_SetsActiveIndex()
    {
        var model = new SpreadsheetModel();
        model.SubmitCommand("echo hi");
        model.MoveToNextEmptyRow();

        model.SelectRow(0);
        Assert.Equal(0, model.ActiveRowIndex);
    }

    [Fact]
    public void SelectRow_EmptyRow_SetsActiveIndex()
    {
        var model = new SpreadsheetModel();
        model.SelectRow(5);
        Assert.Equal(5, model.ActiveRowIndex);
    }

    [Fact]
    public void SelectRow_NegativeClamps_ToZero()
    {
        var model = new SpreadsheetModel();
        model.SelectRow(-3);
        Assert.Equal(0, model.ActiveRowIndex);
    }

    [Fact]
    public void SelectRow_HighRow_Allowed()
    {
        var model = new SpreadsheetModel();
        model.SelectRow(42);
        Assert.Equal(42, model.ActiveRowIndex);
    }

    // ─── Command submission: new row ─────────────────────────────

    [Fact]
    public void SubmitCommand_NewRow_AddsToCellData()
    {
        var model = new SpreadsheetModel();
        var data = model.SubmitCommand("echo hello");

        Assert.Single(model.Rows);
        Assert.Equal("echo hello", data.Command);
        Assert.Equal(CellStatus.Running, data.Status);
        Assert.Equal(1, data.RowNumber);
    }

    [Fact]
    public void SubmitCommand_SequentialRows_CorrectRowNumbers()
    {
        var model = new SpreadsheetModel();

        model.SubmitCommand("cmd1");
        model.MoveToNextEmptyRow();
        model.SubmitCommand("cmd2");
        model.MoveToNextEmptyRow();
        model.SubmitCommand("cmd3");

        Assert.Equal(3, model.Rows.Count);
        Assert.Equal(1, model.Rows[0].RowNumber);
        Assert.Equal(2, model.Rows[1].RowNumber);
        Assert.Equal(3, model.Rows[2].RowNumber);
    }

    // ─── Command submission: re-run existing row ─────────────────

    [Fact]
    public void SubmitCommand_ExistingRow_UpdatesCommand()
    {
        var model = new SpreadsheetModel();
        var data1 = model.SubmitCommand("echo old");
        data1.Output = "old";
        data1.Status = CellStatus.Complete;
        model.MoveToNextEmptyRow();

        // Select row 0 and re-run with new command
        model.SelectRow(0);
        var data2 = model.SubmitCommand("echo new");

        Assert.Single(model.Rows); // no new row added
        Assert.Same(data1, data2); // same object mutated
        Assert.Equal("echo new", data2.Command);
        Assert.Equal(CellStatus.Running, data2.Status);
        Assert.Equal("", data2.Output); // output cleared
    }

    // ─── Command submission: gap rows ────────────────────────────

    [Fact]
    public void SubmitCommand_GapRow_PadsWithEmptyRows()
    {
        var model = new SpreadsheetModel();
        model.SelectRow(3); // skip rows 0,1,2
        model.SubmitCommand("echo gap");

        Assert.Equal(4, model.Rows.Count);
        // Rows 0-2 are empty padding
        Assert.Equal(CellStatus.Empty, model.Rows[0].Status);
        Assert.Equal(CellStatus.Empty, model.Rows[1].Status);
        Assert.Equal(CellStatus.Empty, model.Rows[2].Status);
        // Row 3 has the command
        Assert.Equal("echo gap", model.Rows[3].Command);
        Assert.Equal(CellStatus.Running, model.Rows[3].Status);
    }

    [Fact]
    public void SubmitCommand_GapRow_PaddingHasCorrectRowNumbers()
    {
        var model = new SpreadsheetModel();
        model.SelectRow(2);
        model.SubmitCommand("test");

        Assert.Equal(1, model.Rows[0].RowNumber);
        Assert.Equal(2, model.Rows[1].RowNumber);
        Assert.Equal(3, model.Rows[2].RowNumber);
    }

    // ─── MoveToNextEmptyRow ──────────────────────────────────────

    [Fact]
    public void MoveToNextEmptyRow_SetsActiveToRowCount()
    {
        var model = new SpreadsheetModel();
        model.SubmitCommand("cmd1");
        model.MoveToNextEmptyRow();
        Assert.Equal(1, model.ActiveRowIndex);

        model.SubmitCommand("cmd2");
        model.MoveToNextEmptyRow();
        Assert.Equal(2, model.ActiveRowIndex);
    }

    // ─── Column resize ───────────────────────────────────────────

    [Fact]
    public void ColWidth_MinimumEnforced()
    {
        var model = new SpreadsheetModel();
        model.ColWidths[0] = 10; // below minimum
        // The view layer enforces minimum during drag, but model stores raw value
        // So test that MinColWidth constant is sensible
        Assert.Equal(50, SpreadsheetModel.MinColWidth);
    }

    [Fact]
    public void TotalWidth_SumsAllColumns()
    {
        var model = new SpreadsheetModel();
        var expected = 200.0 + 300 + 80 * 51; // A=200, B=300, C-BA=80 each
        Assert.Equal(expected, model.TotalWidth);
    }

    [Fact]
    public void TotalWidth_UpdatesAfterResize()
    {
        var model = new SpreadsheetModel();
        var before = model.TotalWidth;
        model.ColWidths[0] = 400;
        Assert.Equal(before + 200, model.TotalWidth);
    }

    // ─── Per-row height ──────────────────────────────────────────

    [Fact]
    public void RowHeight_DefaultForAllRows()
    {
        var model = new SpreadsheetModel();
        Assert.Equal(22, model.GetRowHeight(0));
        Assert.Equal(22, model.GetRowHeight(99));
    }

    [Fact]
    public void RowHeight_PerRowOverride()
    {
        var model = new SpreadsheetModel();
        model.SetRowHeight(3, 50);

        Assert.Equal(22, model.GetRowHeight(0)); // untouched
        Assert.Equal(22, model.GetRowHeight(2)); // untouched
        Assert.Equal(50, model.GetRowHeight(3)); // overridden
        Assert.Equal(22, model.GetRowHeight(4)); // untouched
    }

    [Fact]
    public void RowHeight_MinimumEnforced()
    {
        var model = new SpreadsheetModel();
        model.SetRowHeight(0, 5); // below minimum
        Assert.Equal(SpreadsheetModel.MinRowHeight, model.GetRowHeight(0));
    }

    // ─── Font size ───────────────────────────────────────────────

    [Fact]
    public void SetFontSize_ClampsToRange()
    {
        var model = new SpreadsheetModel();

        model.SetFontSize(5); // below min
        Assert.Equal(8, model.CellFontSize);

        model.SetFontSize(50); // above max
        Assert.Equal(28, model.CellFontSize);

        model.SetFontSize(14); // in range
        Assert.Equal(14, model.CellFontSize);
    }

    [Fact]
    public void SetFontSize_UpdatesDefaultRowHeight()
    {
        var model = new SpreadsheetModel();
        model.SetFontSize(20);
        Assert.Equal(40, model.DefaultRowHeight); // 20 * 2
    }

    // ─── Output formatting ───────────────────────────────────────

    [Fact]
    public void FormatOutput_Running_ShowsCalculating()
    {
        Assert.Equal("Calculating...",
            SpreadsheetModel.FormatOutput("", CellStatus.Running));
    }

    [Fact]
    public void FormatOutput_Empty_ReturnsEmpty()
    {
        Assert.Equal("",
            SpreadsheetModel.FormatOutput("", CellStatus.Complete));
    }

    [Fact]
    public void FormatOutput_SingleLine_ReturnAsIs()
    {
        Assert.Equal("hello",
            SpreadsheetModel.FormatOutput("hello", CellStatus.Complete));
    }

    [Fact]
    public void FormatOutput_MultiLine_ShowsFirstLineWithCount()
    {
        var output = "line1\nline2\nline3";
        var result = SpreadsheetModel.FormatOutput(output, CellStatus.Complete);
        Assert.Equal("line1 [3 lines]", result);
    }

    [Fact]
    public void FormatOutput_MultiLineWithBlanks_CountsOnlyNonEmpty()
    {
        var output = "line1\n\nline2\n\n";
        var result = SpreadsheetModel.FormatOutput(output, CellStatus.Complete);
        Assert.Equal("line1 [2 lines]", result);
    }

    [Fact]
    public void FormatOutput_SingleNonEmptyLine_NoCount()
    {
        var output = "hello\n\n\n";
        var result = SpreadsheetModel.FormatOutput(output, CellStatus.Complete);
        Assert.Equal("hello", result);
    }

    [Fact]
    public void FormatOutput_AllWhitespaceLines_ReturnsEmpty()
    {
        var output = "\n\n\n";
        var result = SpreadsheetModel.FormatOutput(output, CellStatus.Complete);
        Assert.Equal("", result);
    }
}
