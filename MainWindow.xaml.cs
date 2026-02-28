using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CellShell.Core;

namespace CellShell;

public partial class MainWindow : Window
{
    private static readonly Brush ExcelGreenBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x73, 0x46));
    private static readonly Brush GridLineBrush = new SolidColorBrush(Color.FromRgb(0xD6, 0xD6, 0xD6));
    private static readonly Brush HeaderBgBrush = new SolidColorBrush(Color.FromRgb(0xE6, 0xE6, 0xE6));
    private static readonly Brush ActiveBorderBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x73, 0x46));
    private static readonly FontFamily CellFont = new("Segoe UI");

    static MainWindow()
    {
        ExcelGreenBrush.Freeze();
        GridLineBrush.Freeze();
        HeaderBgBrush.Freeze();
        ActiveBorderBrush.Freeze();
    }

    // ─── Spreadsheet model (public for testing) ──────────────────

    internal readonly SpreadsheetModel Model = new();

    // ─── Drag state ──────────────────────────────────────────────

    private enum DragMode { None, ColumnResize, RowResize }
    private DragMode _dragMode = DragMode.None;
    private int _dragIndex;
    private double _dragStartPos;
    private double _dragStartSize;

    // ─── UI state ────────────────────────────────────────────────

    private TextBox? _activeTextBox;
    private bool _isExecuting;
    private bool _isRedrawing;
    private int _drawnRows = SpreadsheetModel.MinVisibleRows;

    public MainWindow()
    {
        InitializeComponent();

        var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cellshell-debug.log");
        Trace.Listeners.Add(new TextWriterTraceListener(logPath));
        Trace.AutoFlush = true;
        Log("MainWindow init");

        PreviewMouseMove += Window_PreviewMouseMove;
        PreviewMouseUp += Window_PreviewMouseUp;

        Loaded += (_, _) => RedrawAll();
        SizeChanged += (_, _) =>
        {
            var oldCount = Model.ColCount;
            Model.EnsureColumnsForWidth(GridScroller.ViewportWidth);
            if (Model.ColCount != oldCount)
                RedrawAll();
        };
    }

    private static void Log(string msg) =>
        Trace.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {msg}");

    // ─── Top-level redraw ────────────────────────────────────────

    private void RedrawAll()
    {
        // Ensure we always have enough columns to fill the viewport
        Model.EnsureColumnsForWidth(GridScroller.ViewportWidth);

        RedrawHeaders();
        RedrawGrid();
        UpdateNameBox();
        UpdateFormulaBar(Model.ActiveRowIndex < Model.Rows.Count
            ? Model.Rows[Model.ActiveRowIndex].Command : "");
    }

    // ─── Column Headers ────────────────────────────────────────────

    private void RedrawHeaders()
    {
        HeaderCanvas.Children.Clear();
        var headerH = Model.DefaultRowHeight;
        CornerCell.Height = headerH;
        HeaderCanvas.Height = headerH;

        double x = 0;
        for (int i = 0; i < Model.ColCount; i++)
        {
            var colIndex = i;
            var w = Model.ColWidths[i];

            var header = new Border
            {
                Width = w,
                Height = headerH,
                Background = HeaderBgBrush,
                BorderBrush = GridLineBrush,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child = new TextBlock
                {
                    Text = SpreadsheetModel.GetColLetter(i),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontFamily = CellFont,
                    FontSize = Model.CellFontSize,
                    Foreground = Brushes.Black
                }
            };

            var resizeHandle = new Border
            {
                Width = 5,
                HorizontalAlignment = HorizontalAlignment.Right,
                Background = Brushes.Transparent,
                Cursor = Cursors.SizeWE
            };
            resizeHandle.MouseLeftButtonDown += (s, e) =>
            {
                _dragMode = DragMode.ColumnResize;
                _dragIndex = colIndex;
                _dragStartPos = e.GetPosition(this).X;
                _dragStartSize = Model.ColWidths[colIndex];
                Mouse.Capture(this, CaptureMode.SubTree);
                e.Handled = true;
            };

            var container = new Grid { Width = w, Height = headerH };
            container.Children.Add(header);
            container.Children.Add(resizeHandle);

            Canvas.SetLeft(container, x);
            Canvas.SetTop(container, 0);
            HeaderCanvas.Children.Add(container);

            x += w;
        }

        HeaderCanvas.Width = x;
    }

    // ─── Grid Drawing ─────────────────────────────────────────────

    private void RedrawGrid()
    {
        var sw = Stopwatch.StartNew();
        CellCanvas.Children.Clear();
        RowNumbersPanel.Children.Clear();

        _drawnRows = Math.Max(_drawnRows, Math.Max(Model.Rows.Count + 1, SpreadsheetModel.MinVisibleRows));
        var totalRows = _drawnRows;
        CellCanvas.Width = Model.TotalWidth;

        double y = 0;
        for (int r = 0; r < totalRows; r++)
        {
            var rh = Model.GetRowHeight(r);

            // Row number
            var rowNumBorder = new Border
            {
                Width = SpreadsheetModel.RowNumberWidth,
                Height = rh,
                Background = HeaderBgBrush,
                BorderBrush = GridLineBrush,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child = new TextBlock
                {
                    Text = (r + 1).ToString(),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontFamily = CellFont,
                    FontSize = Math.Max(8, Model.CellFontSize - 1),
                    Foreground = Brushes.Black
                }
            };

            var rowIndex = r;
            var rowResizeHandle = new Border
            {
                Height = 3,
                VerticalAlignment = VerticalAlignment.Bottom,
                Background = Brushes.Transparent,
                Cursor = Cursors.SizeNS
            };
            rowResizeHandle.MouseLeftButtonDown += (s, e) =>
            {
                _dragMode = DragMode.RowResize;
                _dragIndex = rowIndex;
                _dragStartPos = e.GetPosition(this).Y;
                _dragStartSize = Model.GetRowHeight(rowIndex);
                Mouse.Capture(this, CaptureMode.SubTree);
                e.Handled = true;
            };

            var rowNumContainer = new Grid
            {
                Width = SpreadsheetModel.RowNumberWidth,
                Height = rh
            };
            rowNumContainer.Children.Add(rowNumBorder);
            rowNumContainer.Children.Add(rowResizeHandle);
            RowNumbersPanel.Children.Add(rowNumContainer);

            // Cells
            double x = 0;
            for (int c = 0; c < Model.ColCount; c++)
            {
                var isActiveCell = r == Model.ActiveRowIndex && c == 0;
                var borderThickness = isActiveCell
                    ? new Thickness(2)
                    : new Thickness(0, 0, 1, 1);
                var borderBrush = isActiveCell ? ActiveBorderBrush : GridLineBrush;

                var cellBorder = new Border
                {
                    Width = Model.ColWidths[c],
                    Height = rh,
                    Background = Brushes.White,
                    BorderBrush = borderBrush,
                    BorderThickness = borderThickness
                };

                // Click to select any row
                var clickTarget = r;
                if (!isActiveCell)
                {
                    cellBorder.Cursor = Cursors.Arrow;
                    cellBorder.MouseLeftButtonDown += (s, e) =>
                    {
                        if (_isExecuting || _dragMode != DragMode.None) return;
                        Log($"Cell click: row={clickTarget}");
                        Model.SelectRow(clickTarget);
                        RedrawAll();
                        e.Handled = true;
                    };
                }

                // Active row
                if (r == Model.ActiveRowIndex)
                {
                    if (c == 0)
                    {
                        var prefill = r < Model.Rows.Count ? Model.Rows[r].Command : "";
                        var tb = new TextBox
                        {
                            Text = prefill,
                            BorderThickness = new Thickness(0),
                            FontFamily = CellFont,
                            FontSize = Model.CellFontSize,
                            VerticalContentAlignment = VerticalAlignment.Center,
                            Padding = new Thickness(2, 0, 0, 0),
                            Background = Brushes.Transparent
                        };
                        tb.KeyDown += ActiveCell_KeyDown;
                        tb.TextChanged += ActiveCell_TextChanged;
                        cellBorder.Child = tb;
                        _activeTextBox = tb;
                    }
                    else if (c == 1 && r < Model.Rows.Count)
                    {
                        var output = SpreadsheetModel.FormatOutput(
                            Model.Rows[r].Output, Model.Rows[r].Status);
                        cellBorder.Child = new TextBlock
                        {
                            Text = output,
                            VerticalAlignment = VerticalAlignment.Center,
                            Padding = new Thickness(4, 0, 0, 0),
                            FontFamily = CellFont,
                            FontSize = Model.CellFontSize,
                            TextTrimming = TextTrimming.CharacterEllipsis,
                            Foreground = Model.Rows[r].Status == CellStatus.Error
                                ? Brushes.Red : Brushes.Black
                        };
                    }
                }
                else if (r < Model.Rows.Count)
                {
                    var data = Model.Rows[r];
                    if (c == 0)
                    {
                        cellBorder.Child = new TextBlock
                        {
                            Text = data.Command,
                            VerticalAlignment = VerticalAlignment.Center,
                            Padding = new Thickness(4, 0, 0, 0),
                            FontFamily = CellFont,
                            FontSize = Model.CellFontSize,
                            TextTrimming = TextTrimming.CharacterEllipsis
                        };
                    }
                    else if (c == 1)
                    {
                        var output = SpreadsheetModel.FormatOutput(data.Output, data.Status);
                        cellBorder.Child = new TextBlock
                        {
                            Text = output,
                            VerticalAlignment = VerticalAlignment.Center,
                            Padding = new Thickness(4, 0, 0, 0),
                            FontFamily = CellFont,
                            FontSize = Model.CellFontSize,
                            TextTrimming = TextTrimming.CharacterEllipsis,
                            Foreground = data.Status == CellStatus.Error
                                ? Brushes.Red : Brushes.Black
                        };
                    }
                }

                Canvas.SetLeft(cellBorder, x);
                Canvas.SetTop(cellBorder, y);
                CellCanvas.Children.Add(cellBorder);

                x += Model.ColWidths[c];
            }

            y += rh;
        }

        CellCanvas.Height = y;
        sw.Stop();
        Log($"RedrawGrid: {totalRows} rows, {CellCanvas.Children.Count} cells, {sw.ElapsedMilliseconds}ms, active={Model.ActiveRowIndex}");

        var tbToFocus = _activeTextBox;
        if (tbToFocus != null)
        {
            Dispatcher.BeginInvoke(() =>
            {
                tbToFocus.Focus();
                tbToFocus.CaretIndex = tbToFocus.Text.Length;
            });
        }
    }

    // ─── UI updates ──────────────────────────────────────────────

    private void UpdateNameBox()
    {
        NameBox.Text = $"A{Model.ActiveRowIndex + 1}";
    }

    private void UpdateFormulaBar(string command)
    {
        var escaped = command.Replace("\"", "\"\"");
        FormulaText.Text = $"=EXEC(\"{escaped}\")";
    }

    // ─── Window-level drag ───────────────────────────────────────

    private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (_dragMode == DragMode.ColumnResize)
        {
            var delta = e.GetPosition(this).X - _dragStartPos;
            Model.ColWidths[_dragIndex] = Math.Max(
                SpreadsheetModel.MinColWidth, _dragStartSize + delta);
            RedrawHeaders();
            e.Handled = true;
        }
        else if (_dragMode == DragMode.RowResize)
        {
            var delta = e.GetPosition(this).Y - _dragStartPos;
            Model.SetRowHeight(_dragIndex,
                Math.Max(SpreadsheetModel.MinRowHeight, (int)(_dragStartSize + delta)));
            e.Handled = true;
        }
    }

    private void Window_PreviewMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_dragMode != DragMode.None)
        {
            Log($"Drag end: {_dragMode} index={_dragIndex}");
            _dragMode = DragMode.None;
            Mouse.Capture(null);
            RedrawAll();
            e.Handled = true;
        }
    }

    // ─── Cell events ─────────────────────────────────────────────

    private void ActiveCell_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox tb)
            UpdateFormulaBar(tb.Text);
    }

    private async void ActiveCell_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter) return;
        if (sender is not TextBox tb) return;
        if (_isExecuting) return;

        var command = tb.Text.Trim();
        if (string.IsNullOrEmpty(command)) return;

        e.Handled = true;
        _isExecuting = true;
        tb.KeyDown -= ActiveCell_KeyDown;
        tb.TextChanged -= ActiveCell_TextChanged;
        tb.IsReadOnly = true;

        var data = Model.SubmitCommand(command);

        StatusText.Text = " Calculating...";
        Log($"Execute: '{command}' (row {Model.ActiveRowIndex})");

        try
        {
            try
            {
                var output = await CommandExecutor.ExecuteAsync(command);
                data.Output = output;
                data.Status = CellStatus.Complete;
            }
            catch (Exception ex)
            {
                data.Output = ex.Message;
                data.Status = CellStatus.Error;
            }

            StatusText.Text = " Ready";
            Title = $"{CommandExecutor.WorkingDirectory} - Excel";
            Log($"Execute done: status={data.Status}, output={data.Output.Length} chars");

            Model.MoveToNextEmptyRow();
            RedrawAll();
            ScrollToActiveRow();
        }
        finally
        {
            _isExecuting = false;
        }
    }

    private void ScrollToActiveRow()
    {
        double targetY = 0;
        for (int i = 0; i < Model.ActiveRowIndex; i++)
            targetY += Model.GetRowHeight(i);
        var rowH = Model.GetRowHeight(Model.ActiveRowIndex);
        GridScroller.ScrollToVerticalOffset(
            Math.Max(0, targetY - GridScroller.ViewportHeight + rowH * 2));
    }

    // ─── Scroll sync ─────────────────────────────────────────────

    private void GridScroller_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        RowNumberScroller.ScrollToVerticalOffset(e.VerticalOffset);
        HeaderCanvas.RenderTransform = new TranslateTransform(-e.HorizontalOffset, 0);

        // Infinite scroll: extend grid when user scrolls near the bottom
        // Only extend when user has actually scrolled (VerticalOffset > 0) to prevent
        // startup cascade — layout changes fire ScrollChanged which would loop forever
        if (!_isRedrawing && e.VerticalOffset > 0
            && e.VerticalOffset + e.ViewportHeight >= e.ExtentHeight - Model.DefaultRowHeight * 5)
        {
            _isRedrawing = true;
            _drawnRows += 30;
            RedrawGrid();
            _isRedrawing = false;
        }
    }

    // ─── Text size ───────────────────────────────────────────────

    private void Menu_TextSizeUp(object sender, RoutedEventArgs e)
    {
        Model.SetFontSize(Model.CellFontSize + 1);
        RedrawAll();
    }

    private void Menu_TextSizeDown(object sender, RoutedEventArgs e)
    {
        Model.SetFontSize(Model.CellFontSize - 1);
        RedrawAll();
    }

    // ─── Joke Menu Handlers ───────────────────────────────────────

    private void Menu_NewWindow(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Sure, let me open another\nterminal disguised as Excel.\n\nJust kidding, use Alt+Enter.", "New Window", MessageBoxButton.OK);
    private void Menu_Open(object sender, RoutedEventArgs e) =>
        MessageBox.Show("What would you even open?\nThis is a terminal.", "Open", MessageBoxButton.OK);
    private void Menu_Save(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Cannot save terminal session\nas a spreadsheet.\n\nNice try though.", "Save", MessageBoxButton.OK);
    private void Menu_Exit(object sender, RoutedEventArgs e) => Close();
    private void Menu_Bold(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Everything is already bold.\nIt's a terminal.", "Bold", MessageBoxButton.OK);
    private void Menu_Italic(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Monospace fonts don't do\nitalic. Sorry.", "Italic", MessageBoxButton.OK);
    private void Menu_FontSize(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Use View > Text Size +/- to actually\nchange the font size.\n\nBet you didn't expect that to work.", "Font Size", MessageBoxButton.OK);
    private void Menu_InsertFunction(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Available functions:\n\n=EXEC(command)\n  Executes a shell command\n\nThat's it. That's the whole list.", "Insert Function", MessageBoxButton.OK);
    private void Menu_InsertChart(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Have you tried:\n  echo '|||||||'\n  echo '|||||'\n  echo '|||'\n  echo '|'", "Insert Chart", MessageBoxButton.OK);
    private void Menu_Margins(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Your margins are exactly\n0 pixels on all sides.\n\nMaximum productivity.", "Margins", MessageBoxButton.OK);
    private void Menu_Orientation(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Landscape? Portrait?\nIt's a terminal.\nIt's whatever shape your\nwindow is.", "Orientation", MessageBoxButton.OK);
    private void Menu_PrintArea(object sender, RoutedEventArgs e) =>
        MessageBox.Show("The print area is your\nentire screen.\n\nJust screenshot it.", "Print Area", MessageBoxButton.OK);
    private void Menu_RefreshAll(object sender, RoutedEventArgs e)
    {
        if (Model.Rows.Count == 0)
        {
            MessageBox.Show("Nothing to refresh.\nType some commands first.", "Refresh", MessageBoxButton.OK);
            return;
        }
        var count = Model.Rows.Count(r => r.Status == CellStatus.Complete);
        MessageBox.Show($"Re-running {count} command(s)?\n\nThat sounds dangerous.\nI'll pass.", "Refresh All", MessageBoxButton.OK);
    }
    private void Menu_Sort(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Commands are sorted\nchronologically.\n\nThat's called 'history'.", "Sort", MessageBoxButton.OK);
    private void Menu_SpellCheck(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Spell checking your\ncommands would imply\nI care about your typos.\n\nI don't.", "Spell Check", MessageBoxButton.OK);
    private void Menu_TrackChanges(object sender, RoutedEventArgs e) =>
        MessageBox.Show("All changes are tracked.\nIt's called 'history'.\n\nTry the 'history' command.", "Track Changes", MessageBoxButton.OK);
    private void Menu_FormulaBar(object sender, RoutedEventArgs e) =>
        MessageBox.Show("The formula bar is always visible.\nYou need =EXEC() in your life.", "Formula Bar", MessageBoxButton.OK);
    private void Menu_ZoomIn(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Zooming changes nothing.\nIt's a terminal.", "Zoom", MessageBoxButton.OK);
    private void Menu_ZoomOut(object sender, RoutedEventArgs e) =>
        MessageBox.Show("Still changes nothing.\nMonospace is monospace.", "Zoom", MessageBoxButton.OK);
    private void Menu_About(object sender, RoutedEventArgs e) =>
        MessageBox.Show("CellShell v1.0\n\nA terminal disguised as Excel.\nBecause why not.\n\nNobody asked for this.", "About CellShell", MessageBoxButton.OK);
    private void Menu_Clippy(object sender, RoutedEventArgs e)
    {
        var tips = new[]
        {
            "It looks like you're trying to\nuse a terminal. Would you like\nhelp with that?",
            "Did you know? 'rm -rf /' is\nNOT a valid Excel formula.",
            "Pro tip: Ctrl+C doesn't copy\nin a terminal. It kills things.",
            "I see you're typing commands.\nHave you tried turning it off\nand on again?",
            "Fun fact: This spreadsheet has\nexactly 0 formulas and\ninfinite rows.",
            "Tip: Use 'cls' to clear the\nscreen. Just kidding, that\nwon't work here.",
            "Remember: In Excel, every\nproblem is a VLOOKUP away.\nHere, every problem is a\n'command not found' away."
        };
        MessageBox.Show(tips[Random.Shared.Next(tips.Length)], "Clippy", MessageBoxButton.OK);
    }
}
