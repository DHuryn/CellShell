namespace CellShell.Core;

public enum CellStatus
{
    Empty,
    Running,
    Complete,
    Error
}

public class CellData
{
    public int RowNumber { get; set; }
    public string Command { get; set; } = string.Empty;
    public string Output { get; set; } = string.Empty;
    public CellStatus Status { get; set; } = CellStatus.Empty;
}
