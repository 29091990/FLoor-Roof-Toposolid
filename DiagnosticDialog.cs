using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace FloorRoofTopo
{
    /// <summary>
    /// Dark-themed WPF diagnostic dialog showing detailed conversion analysis.
    /// Displays shape point data, Z offsets, and conversion results in a table.
    /// Built entirely in C# code (no XAML required).
    /// </summary>
    public class DiagnosticDialog : Window
    {
        // --- Colors (consistent with ConvertDialog dark theme) ---
        private static readonly Color BgDark = Color.FromRgb(28, 28, 32);
        private static readonly Color BgCard = Color.FromRgb(38, 38, 44);
        private static readonly Color BgTableEven = Color.FromRgb(32, 32, 38);
        private static readonly Color BgTableOdd = Color.FromRgb(38, 38, 44);
        private static readonly Color BgTableHeader = Color.FromRgb(25, 25, 30);
        private static readonly Color AccentBlue = Color.FromRgb(56, 152, 255);
        private static readonly Color AccentGreen = Color.FromRgb(72, 199, 142);
        private static readonly Color AccentRed = Color.FromRgb(232, 72, 85);
        private static readonly Color AccentYellow = Color.FromRgb(255, 193, 7);
        private static readonly Color TextPrimary = Color.FromRgb(235, 235, 240);
        private static readonly Color TextSecondary = Color.FromRgb(160, 160, 175);
        private static readonly Color TextDim = Color.FromRgb(120, 120, 135);
        private static readonly Color BorderColor = Color.FromRgb(55, 55, 65);
        private static readonly Color CardHighlight = Color.FromRgb(45, 48, 58);

        // --- Unit conversion ---
        private const double FT_TO_M = 0.3048;
        private const double FT_TO_MM = 304.8;

        public DiagnosticDialog(
            List<ConversionDiagnosticInfo> diagnostics,
            int successCount, int skippedCount, int failedCount)
        {
            Title = "Phân tích chuyển đổi";
            Width = 800;
            Height = 660;
            MinWidth = 650;
            MinHeight = 450;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;

            // Main border with rounded corners and shadow
            var mainBorder = new Border
            {
                Background = new SolidColorBrush(BgDark),
                CornerRadius = new CornerRadius(12),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    BlurRadius = 30,
                    Opacity = 0.6,
                    ShadowDepth = 5
                }
            };

            var rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(
                new RowDefinition { Height = new GridLength(36) });     // Title bar
            rootGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });        // Header
            rootGrid.RowDefinitions.Add(
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            rootGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });        // Footer

            // === Title Bar ===
            var titleBar = CreateTitleBar();
            Grid.SetRow(titleBar, 0);
            rootGrid.Children.Add(titleBar);

            // === Header with summary ===
            var header = CreateHeader(successCount, skippedCount, failedCount,
                diagnostics?.Count ?? 0);
            Grid.SetRow(header, 1);
            rootGrid.Children.Add(header);

            // === Scrollable content ===
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(0)
            };

            var contentStack = new StackPanel
            {
                Margin = new Thickness(20, 8, 20, 12)
            };

            if (diagnostics != null)
            {
                int cardIdx = 0;
                foreach (var diag in diagnostics)
                {
                    contentStack.Children.Add(CreateElementCard(diag, cardIdx++));
                }
            }

            scrollViewer.Content = contentStack;
            Grid.SetRow(scrollViewer, 2);
            rootGrid.Children.Add(scrollViewer);

            // === Footer ===
            var footer = CreateFooter();
            Grid.SetRow(footer, 3);
            rootGrid.Children.Add(footer);

            mainBorder.Child = rootGrid;
            Content = mainBorder;
        }

        // ============================================================
        //  UI BUILDING BLOCKS
        // ============================================================

        private Border CreateTitleBar()
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = GridLength.Auto });

            var titleText = new TextBlock
            {
                Text = "  📊  PHÂN TÍCH CHUYỂN ĐỔI",
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(TextSecondary),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };
            Grid.SetColumn(titleText, 0);
            grid.Children.Add(titleText);

            var closeBtn = new Button
            {
                Content = "✕",
                Width = 36,
                Height = 28,
                FontSize = 12,
                Background = Brushes.Transparent,
                Foreground = new SolidColorBrush(TextSecondary),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            closeBtn.Click += (s, e) => Close();
            Grid.SetColumn(closeBtn, 1);
            grid.Children.Add(closeBtn);

            var border = new Border
            {
                Height = 36,
                Background = new SolidColorBrush(Color.FromRgb(22, 22, 26)),
                CornerRadius = new CornerRadius(12, 12, 0, 0),
                Child = grid
            };
            border.MouseLeftButtonDown += (s, e) =>
            {
                if (e.LeftButton == MouseButtonState.Pressed) DragMove();
            };

            return border;
        }

        private StackPanel CreateHeader(int success, int skipped,
            int failed, int totalElements)
        {
            var sp = new StackPanel { Margin = new Thickness(24, 14, 24, 6) };

            // Title
            var title = new TextBlock
            {
                Text = "Kết quả chuyển đổi",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new LinearGradientBrush(AccentBlue, AccentGreen, 0)
            };
            sp.Children.Add(title);

            // Summary stats
            var statsPanel = new WrapPanel
            {
                Margin = new Thickness(0, 8, 0, 0)
            };

            if (success > 0)
                statsPanel.Children.Add(CreateStatBadge(
                    $"✅ {success} thành công", AccentGreen));
            if (skipped > 0)
                statsPanel.Children.Add(CreateStatBadge(
                    $"⏭ {skipped} bỏ qua", AccentYellow));
            if (failed > 0)
                statsPanel.Children.Add(CreateStatBadge(
                    $"❌ {failed} thất bại", AccentRed));

            sp.Children.Add(statsPanel);

            // Separator
            sp.Children.Add(new Border
            {
                Background = new SolidColorBrush(BorderColor),
                Height = 1,
                Margin = new Thickness(0, 10, 0, 0)
            });

            return sp;
        }

        private Border CreateStatBadge(string text, Color color)
        {
            return new Border
            {
                Background = new SolidColorBrush(
                    Color.FromArgb(30, color.R, color.G, color.B)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 4, 10, 4),
                Margin = new Thickness(0, 0, 8, 0),
                Child = new TextBlock
                {
                    Text = text,
                    FontSize = 12,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(color)
                }
            };
        }

        /// <summary>
        /// Create a card showing detailed conversion analysis for one element.
        /// </summary>
        private Border CreateElementCard(ConversionDiagnosticInfo diag, int index)
        {
            var stack = new StackPanel();

            // ── Card header ───────────────────────────────────────────
            var headerPanel = new DockPanel
            {
                Margin = new Thickness(0, 0, 0, 8)
            };

            Color statusColor = diag.Success ? AccentGreen : AccentRed;
            string statusIcon = diag.Success ? "✅" : "❌";

            var headerLeft = new TextBlock
            {
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(TextPrimary)
            };
            headerLeft.Inlines.Add(new System.Windows.Documents.Run(
                $"{statusIcon}  {diag.SourceType} → {diag.TargetType}")
            {
                Foreground = new SolidColorBrush(TextPrimary)
            });
            DockPanel.SetDock(headerLeft, Dock.Left);
            headerPanel.Children.Add(headerLeft);

            var headerRight = new TextBlock
            {
                Text = $"ID: {diag.SourceId}",
                FontSize = 11,
                Foreground = new SolidColorBrush(TextDim),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(headerRight);

            stack.Children.Add(headerPanel);

            // ── Element info grid ─────────────────────────────────────
            var infoGrid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 8)
            };
            infoGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            infoGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left: Source info
            var srcInfo = new StackPanel();
            srcInfo.Children.Add(CreateInfoLabel("NGUỒN", AccentBlue));
            srcInfo.Children.Add(CreateInfoLine(
                $"Level: {diag.LevelName} ({FormatM(diag.LevelElevation)}m)"));
            srcInfo.Children.Add(CreateInfoLine(
                $"Offset: {FormatMM(diag.SourceOffset)}mm"));
            srcInfo.Children.Add(CreateInfoLine(
                $"Base Z: {FormatM(diag.SourceBaseZ)}m"));
            srcInfo.Children.Add(CreateInfoLine(
                $"Biên: {diag.BoundaryLoops} loop, {diag.TotalCurves} cạnh"));

            Grid.SetColumn(srcInfo, 0);
            infoGrid.Children.Add(srcInfo);

            // Right: Target info
            var tgtInfo = new StackPanel();
            tgtInfo.Children.Add(CreateInfoLabel("ĐÍCH", AccentGreen));
            tgtInfo.Children.Add(CreateInfoLine(
                $"Loại: {diag.TargetType} (ID: {diag.TargetId})"));
            tgtInfo.Children.Add(CreateInfoLine(
                $"Base Z: {FormatM(diag.TargetBaseZ)}m"));
            tgtInfo.Children.Add(CreateInfoLine(
                $"Điểm: {diag.AppliedCount}/{diag.ShapePoints.Count} áp dụng"));

            if (!string.IsNullOrEmpty(diag.Message) && !diag.Success)
            {
                tgtInfo.Children.Add(new TextBlock
                {
                    Text = diag.Message,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(AccentRed),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 2, 0, 0)
                });
            }

            Grid.SetColumn(tgtInfo, 1);
            infoGrid.Children.Add(tgtInfo);

            stack.Children.Add(infoGrid);

            // ── Shape points table ────────────────────────────────────
            if (diag.ShapePoints != null && diag.ShapePoints.Count > 0)
            {
                stack.Children.Add(CreateInfoLabel(
                    $"ĐIỂM NHẤC ({diag.ShapePoints.Count} điểm)", AccentBlue));
                stack.Children.Add(CreatePointsTable(diag.ShapePoints));
            }
            else
            {
                stack.Children.Add(new TextBlock
                {
                    Text = "(Không có điểm nhấc SlabShapeEditor)",
                    FontSize = 11,
                    Foreground = new SolidColorBrush(TextDim),
                    FontStyle = FontStyles.Italic,
                    Margin = new Thickness(0, 4, 0, 0)
                });
            }

            // ── Wrap in card border ───────────────────────────────────
            return new Border
            {
                Background = new SolidColorBrush(BgCard),
                CornerRadius = new CornerRadius(10),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(16, 12, 16, 12),
                Margin = new Thickness(0, 4, 0, 4),
                Child = stack
            };
        }

        /// <summary>
        /// Create the shape points data table using a WPF Grid control.
        /// Columns: # | Loại | X(m) | Y(m) | Z nguồn(m) | ΔZ(mm) | Trạng thái
        /// </summary>
        private UIElement CreatePointsTable(List<ShapePointData> points)
        {
            var tableGrid = new Grid
            {
                Margin = new Thickness(0, 4, 0, 0)
            };

            // Column widths
            double[] colWidths = { 35, 65, 90, 90, 90, 80, 80, 40 };
            foreach (var w in colWidths)
            {
                tableGrid.ColumnDefinitions.Add(
                    new ColumnDefinition { Width = new GridLength(w) });
            }

            // Header row
            tableGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(26) });
            string[] headers = { "#", "Loại", "X (m)", "Y (m)",
                "Z nguồn", "ΔZ (mm)", "Z đích", "OK" };

            var headerBg = new Border
            {
                Background = new SolidColorBrush(BgTableHeader),
                CornerRadius = new CornerRadius(4, 4, 0, 0)
            };
            Grid.SetRow(headerBg, 0);
            Grid.SetColumnSpan(headerBg, headers.Length);
            tableGrid.Children.Add(headerBg);

            for (int c = 0; c < headers.Length; c++)
            {
                var tb = new TextBlock
                {
                    Text = headers[c],
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(AccentBlue),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 4, 0)
                };
                Grid.SetRow(tb, 0);
                Grid.SetColumn(tb, c);
                tableGrid.Children.Add(tb);
            }

            // Data rows
            for (int i = 0; i < points.Count; i++)
            {
                var pt = points[i];
                int row = i + 1;
                tableGrid.RowDefinitions.Add(
                    new RowDefinition { Height = new GridLength(22) });

                // Alternating row background
                Color rowColor = i % 2 == 0 ? BgTableEven : BgTableOdd;
                var rowBg = new Border
                {
                    Background = new SolidColorBrush(rowColor)
                };
                if (i == points.Count - 1) // Last row: rounded bottom
                    rowBg.CornerRadius = new CornerRadius(0, 0, 4, 4);
                Grid.SetRow(rowBg, row);
                Grid.SetColumnSpan(rowBg, headers.Length);
                tableGrid.Children.Add(rowBg);

                // Highlight rows with non-zero ΔZ
                bool hasOffset = Math.Abs(pt.RelativeZ) > 1e-6;
                Color textColor = hasOffset ? AccentYellow : TextPrimary;

                AddTableCell(tableGrid, row, 0,
                    (pt.Index + 1).ToString(), TextSecondary);
                AddTableCell(tableGrid, row, 1,
                    TranslateVertexType(pt.VertexType), TextSecondary);
                AddTableCell(tableGrid, row, 2,
                    FormatM(pt.X), TextPrimary);
                AddTableCell(tableGrid, row, 3,
                    FormatM(pt.Y), TextPrimary);
                AddTableCell(tableGrid, row, 4,
                    FormatM(pt.SourceZ), TextPrimary);
                AddTableCell(tableGrid, row, 5,
                    FormatMM(pt.RelativeZ), textColor);
                AddTableCell(tableGrid, row, 6,
                    FormatM(pt.TargetZ), TextPrimary);

                // Status indicator
                var statusTB = new TextBlock
                {
                    Text = pt.Applied ? "✓" : "✗",
                    FontSize = 13,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(
                        pt.Applied ? AccentGreen : AccentRed),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                Grid.SetRow(statusTB, row);
                Grid.SetColumn(statusTB, 7);
                tableGrid.Children.Add(statusTB);
            }

            // Wrap in ScrollViewer if too many points
            if (points.Count > 12)
            {
                return new ScrollViewer
                {
                    MaxHeight = 280,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    Content = tableGrid,
                    Margin = new Thickness(0, 0, 0, 4)
                };
            }

            return tableGrid;
        }

        private Grid CreateFooter()
        {
            var grid = new Grid
            {
                Margin = new Thickness(24, 8, 24, 16)
            };
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = GridLength.Auto });

            // Unit note
            var unitNote = new TextBlock
            {
                Text = "Đơn vị: m (mét) / mm (milimet) — Chuyển từ feet nội bộ Revit",
                FontSize = 10,
                Foreground = new SolidColorBrush(TextDim),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(unitNote, 0);
            grid.Children.Add(unitNote);

            // Close button
            var btnClose = new Button
            {
                Content = "  Đóng  ",
                Height = 34,
                MinWidth = 100,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Background = new LinearGradientBrush(
                    AccentBlue, Color.FromRgb(36, 120, 220), 90),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => Close();
            Grid.SetColumn(btnClose, 1);
            grid.Children.Add(btnClose);

            return grid;
        }

        // ============================================================
        //  HELPER METHODS
        // ============================================================

        private void AddTableCell(Grid grid, int row, int col,
            string text, Color foreground)
        {
            var tb = new TextBlock
            {
                Text = text,
                FontSize = 10.5,
                Foreground = new SolidColorBrush(foreground),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 4, 0)
            };
            Grid.SetRow(tb, row);
            Grid.SetColumn(tb, col);
            grid.Children.Add(tb);
        }

        private TextBlock CreateInfoLabel(string text, Color color)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(color),
                Margin = new Thickness(0, 4, 0, 4)
            };
        }

        private TextBlock CreateInfoLine(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 11.5,
                Foreground = new SolidColorBrush(TextPrimary),
                Margin = new Thickness(0, 1, 0, 1)
            };
        }

        private static string FormatM(double feet)
        {
            return (feet * FT_TO_M).ToString("F3");
        }

        private static string FormatMM(double feet)
        {
            return (feet * FT_TO_MM).ToString("F1");
        }

        private static string TranslateVertexType(string type)
        {
            switch (type)
            {
                case "Corner": return "Góc";
                case "Edge": return "Cạnh";
                case "Interior": return "Nội bộ";
                case "ContourBoundary": return "Viền";
                default: return type ?? "?";
            }
        }
    }
}
