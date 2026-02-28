using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
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
    private bool _isRedrawing;
    private int _drawnRows = SpreadsheetModel.MinVisibleRows;
    private Border? _expandedOverlay;
    private int _expandedRow = -1;
    private TextBlock? _expandedTextBlock;
    private ScrollViewer? _expandedScroller;
    private readonly Dictionary<int, TextBlock> _outputCells = new();
    private readonly Dictionary<int, Border> _outputBorders = new();

    public MainWindow()
    {
        InitializeComponent();

        var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cellshell-debug.log");
        Trace.Listeners.Add(new TextWriterTraceListener(logPath));
        Trace.AutoFlush = true;
        Log("MainWindow init");

        PreviewMouseMove += Window_PreviewMouseMove;
        PreviewMouseUp += Window_PreviewMouseUp;
        PreviewMouseLeftButtonDown += (_, e) =>
        {
            if (_expandedOverlay != null && _expandedOverlay.IsMouseOver)
            {
                e.Handled = true;
                return;
            }
            if (e.ClickCount == 2)
                HandleCellDoubleClick(e);
            else
                DismissExpanded();
        };
        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape) DismissExpanded();
            if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
                HandleCtrlC(e);
        };

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
        DismissExpanded();

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
        DismissExpanded();

        var sw = Stopwatch.StartNew();
        CellCanvas.Children.Clear();
        RowNumbersPanel.Children.Clear();
        _outputCells.Clear();
        _outputBorders.Clear();

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
                        if (_dragMode != DragMode.None) return;
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
                        var outputTb = new TextBlock
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
                        cellBorder.Child = outputTb;
                        _outputCells[r] = outputTb;
                        _outputBorders[r] = cellBorder;
                        if (Model.Rows[r].Status == CellStatus.Running)
                            ApplyRunningAnimation(cellBorder);
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
                        var outputTb = new TextBlock
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
                        cellBorder.Child = outputTb;
                        _outputCells[r] = outputTb;
                        _outputBorders[r] = cellBorder;
                        if (data.Status == CellStatus.Running)
                            ApplyRunningAnimation(cellBorder);
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

    private void ApplyRunningAnimation(Border border)
    {
        border.BorderThickness = new Thickness(2);
        var animBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x73, 0x46));
        border.BorderBrush = animBrush;
        var anim = new ColorAnimation
        {
            From = Color.FromRgb(0x21, 0x73, 0x46),
            To = Color.FromRgb(0xA0, 0xD8, 0xB0),
            Duration = TimeSpan.FromSeconds(0.7),
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever
        };
        animBrush.BeginAnimation(SolidColorBrush.ColorProperty, anim);
    }

    private static void StopRunningAnimation(Border border)
    {
        if (border.BorderBrush is SolidColorBrush brush && brush.HasAnimatedProperties)
        {
            brush.BeginAnimation(SolidColorBrush.ColorProperty, null);
        }
        border.BorderBrush = GridLineBrush;
        border.BorderThickness = new Thickness(0, 0, 1, 1);
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
            RedrawAll();
            e.Handled = true;
        }
        else if (_dragMode == DragMode.RowResize)
        {
            var delta = e.GetPosition(this).Y - _dragStartPos;
            Model.SetRowHeight(_dragIndex,
                Math.Max(SpreadsheetModel.MinRowHeight, (int)(_dragStartSize + delta)));
            RedrawAll();
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

    private void HandleCtrlC(KeyEventArgs e)
    {
        // Only kill if the active row is running — don't reach into other rows
        if (Model.ActiveRowIndex >= Model.Rows.Count) return;
        var target = Model.Rows[Model.ActiveRowIndex];
        if (target.Status != CellStatus.Running || target.RunningProcess == null) return;

        e.Handled = true;
        target.WasCancelled = true;
        Log($"Ctrl+C: killing process for row {target.RowNumber}");
        try { target.RunningProcess.Kill(entireProcessTree: true); } catch { }
    }

    private void ActiveCell_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter) return;
        if (sender is not TextBox tb) return;

        var command = tb.Text.Trim();
        if (string.IsNullOrEmpty(command)) return;

        // Guard: don't re-submit if this row is already running
        var currentRow = Model.ActiveRowIndex;
        if (currentRow < Model.Rows.Count && Model.Rows[currentRow].Status == CellStatus.Running)
            return;

        e.Handled = true;
        tb.KeyDown -= ActiveCell_KeyDown;
        tb.TextChanged -= ActiveCell_TextChanged;
        tb.IsReadOnly = true;

        var data = Model.SubmitCommand(command);
        var executingRow = currentRow;

        Log($"Execute: '{command}' (row {executingRow})");

        // Move to next row immediately so user can type the next command
        Model.MoveToNextEmptyRow();
        RedrawAll();
        ScrollToActiveRow();

        // Fire-and-forget the streaming execution
        _ = RunStreamingAsync(data, executingRow, command);
    }

    private async Task RunStreamingAsync(CellData data, int row, string command)
    {
        try
        {
            await CommandExecutor.ExecuteStreamingAsync(
                command,
                line => Dispatcher.InvokeAsync(() =>
                {
                    data.Output = string.IsNullOrEmpty(data.Output)
                        ? line
                        : data.Output + "\n" + line;

                    if (_outputCells.TryGetValue(row, out var tb))
                    {
                        tb.Text = SpreadsheetModel.FormatOutput(data.Output, data.Status);
                    }

                    // Update expanded overlay if it's showing this row
                    if (_expandedRow == row && _expandedTextBlock != null)
                    {
                        _expandedTextBlock.Text = data.Output;
                        _expandedScroller?.ScrollToEnd();
                    }
                }),
                process => Dispatcher.InvokeAsync(() => data.RunningProcess = process));

            Dispatcher.Invoke(() =>
            {
                if (data.WasCancelled)
                {
                    data.Output = string.IsNullOrEmpty(data.Output) ? "^C" : data.Output + "\n^C";
                    data.Status = CellStatus.Error;
                }
                else if (data.Output.Contains("[Timed out"))
                {
                    data.Status = CellStatus.Error;
                }
                else
                {
                    data.Status = CellStatus.Complete;
                }
                data.RunningProcess = null;
                data.WasCancelled = false;

                Title = $"{CommandExecutor.WorkingDirectory} - Excel";
                Log($"Execute done: row={row}, status={data.Status}, output={data.Output.Length} chars");

                if (_outputBorders.TryGetValue(row, out var border))
                    StopRunningAnimation(border);

                if (_outputCells.TryGetValue(row, out var tb))
                {
                    tb.Text = SpreadsheetModel.FormatOutput(data.Output, data.Status);
                    tb.Foreground = data.Status == CellStatus.Error ? Brushes.Red : Brushes.Black;
                }

                FinalizeExpandedOverlay(row, data);
            });
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                data.Output = ex.Message;
                data.Status = CellStatus.Error;
                data.RunningProcess = null;

                Log($"Execute error: row={row}, {ex.Message}");

                if (_outputBorders.TryGetValue(row, out var border))
                    StopRunningAnimation(border);

                if (_outputCells.TryGetValue(row, out var tb))
                {
                    tb.Text = SpreadsheetModel.FormatOutput(data.Output, data.Status);
                    tb.Foreground = Brushes.Red;
                }

                FinalizeExpandedOverlay(row, data);
            });
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

    // ─── Cell expand overlay ──────────────────────────────────────

    private void DismissExpanded()
    {
        if (_expandedOverlay != null)
        {
            CellCanvas.Children.Remove(_expandedOverlay);
            _expandedOverlay = null;
            _expandedRow = -1;
            _expandedTextBlock = null;
            _expandedScroller = null;
        }
    }

    private void HandleCellDoubleClick(MouseButtonEventArgs e)
    {
        var pos = e.GetPosition(CellCanvas);
        if (pos.X < 0 || pos.Y < 0) return;

        // Find row
        int row = -1;
        double cellY = 0;
        for (int r = 0; r < _drawnRows; r++)
        {
            var rh = Model.GetRowHeight(r);
            if (pos.Y < cellY + rh) { row = r; break; }
            cellY += rh;
        }
        if (row < 0 || row >= Model.Rows.Count) return;

        // Find column
        int col = -1;
        double cellX = 0;
        for (int c = 0; c < Model.ColCount; c++)
        {
            var w = Model.ColWidths[c];
            if (pos.X < cellX + w) { col = c; break; }
            cellX += w;
        }
        if (col < 0 || col > 1) return;

        // Get text: column A = command, column B = raw output
        var data = Model.Rows[row];
        string? text = col == 0 ? data.Command : data.Output;
        if (string.IsNullOrEmpty(text)) return;

        // Don't expand the active cell's command (column A) — it's a TextBox for editing
        if (row == Model.ActiveRowIndex && col == 0) return;

        var cellW = Model.ColWidths[col];
        var cellH = Model.GetRowHeight(row);
        bool isErr = col == 1 && data.Status == CellStatus.Error;
        bool isRunning = col == 1 && data.Status == CellStatus.Running;

        ExpandCell(row, text, cellX, cellY, cellW, cellH, isErr, isRunning);
        e.Handled = true;
    }

    private void ExpandCell(int row, string text, double cellLeft, double cellTop,
        double cellWidth, double cellHeight, bool isError, bool isRunning)
    {
        DismissExpanded();

        var maxW = Math.Max(300, GridScroller.ViewportWidth * 0.6);
        var maxH = Math.Max(150, GridScroller.ViewportHeight * 0.6);

        var tb = new TextBlock
        {
            Text = text,
            TextWrapping = TextWrapping.Wrap,
            FontFamily = CellFont,
            FontSize = Model.CellFontSize,
            Padding = new Thickness(6, 4, 6, 4),
            Foreground = isError ? Brushes.Red : Brushes.Black,
        };

        var sv = new ScrollViewer
        {
            Content = tb,
            MaxWidth = maxW,
            MaxHeight = maxH,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        };

        var overlay = new Border
        {
            Background = Brushes.White,
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x40, 0x40, 0x40)),
            Child = sv,
            MinWidth = cellWidth,
            MinHeight = cellHeight
        };

        if (isRunning)
            ApplyRunningAnimation(overlay);

        Canvas.SetLeft(overlay, cellLeft);
        Canvas.SetTop(overlay, cellTop);
        Panel.SetZIndex(overlay, 100);
        CellCanvas.Children.Add(overlay);
        _expandedOverlay = overlay;
        _expandedRow = row;
        _expandedTextBlock = tb;
        _expandedScroller = sv;

        // Auto-scroll to bottom for running cells
        if (isRunning)
            sv.ScrollToEnd();
    }

    private void FinalizeExpandedOverlay(int row, CellData data)
    {
        if (_expandedRow != row || _expandedOverlay == null) return;
        StopRunningAnimation(_expandedOverlay);
        _expandedOverlay.BorderBrush = new SolidColorBrush(Color.FromRgb(0x40, 0x40, 0x40));
        _expandedOverlay.BorderThickness = new Thickness(1);
        if (_expandedTextBlock != null)
        {
            _expandedTextBlock.Text = data.Output;
            _expandedTextBlock.Foreground = data.Status == CellStatus.Error ? Brushes.Red : Brushes.Black;
        }
    }

    // ─── Scroll sync ─────────────────────────────────────────────

    private void GridScroller_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        DismissExpanded();
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
