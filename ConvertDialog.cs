using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;

namespace FloorRoofTopo
{
    /// <summary>
    /// Dark-themed WPF dialog for selecting conversion target type.
    /// Built entirely in C# code (no XAML required).
    /// </summary>
    public class ConvertDialog : Window
    {
        // --- Result Properties ---
        public ConvertTarget SelectedTargetType { get; private set; }
        public bool DeleteSource { get; private set; }

        // --- UI Controls ---
        private RadioButton rbFloor, rbRoof, rbTopo;
        private CheckBox chkDelete;

        // --- Colors ---
        private static readonly Color BgDark = Color.FromRgb(28, 28, 32);
        private static readonly Color BgCard = Color.FromRgb(38, 38, 44);
        private static readonly Color BgHover = Color.FromRgb(50, 50, 58);
        private static readonly Color AccentBlue = Color.FromRgb(56, 152, 255);
        private static readonly Color AccentGreen = Color.FromRgb(72, 199, 142);
        private static readonly Color TextPrimary = Color.FromRgb(235, 235, 240);
        private static readonly Color TextSecondary = Color.FromRgb(160, 160, 175);
        private static readonly Color BorderColor = Color.FromRgb(55, 55, 65);

        public ConvertDialog(string sourceTypes, int count, bool topoAvailable)
        {
            Title = "Floor ↔ Roof ↔ Topo Converter";
            Width = 440;
            Height = 460;
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
                },
                Padding = new Thickness(0)
            };

            var rootStack = new StackPanel();

            // === TITLE BAR (draggable) ===
            var titleBar = CreateTitleBar();
            rootStack.Children.Add(titleBar);

            // === HEADER ===
            var header = CreateHeader();
            rootStack.Children.Add(header);

            // === SOURCE INFO ===
            var sourceInfo = CreateSourceInfo(sourceTypes, count);
            rootStack.Children.Add(sourceInfo);

            // === SEPARATOR ===
            rootStack.Children.Add(CreateSeparator());

            // === TARGET TYPE SECTION ===
            var targetSection = CreateTargetSection(topoAvailable);
            rootStack.Children.Add(targetSection);

            // === SEPARATOR ===
            rootStack.Children.Add(CreateSeparator());

            // === OPTIONS ===
            var options = CreateOptions();
            rootStack.Children.Add(options);

            // === BUTTONS ===
            var buttons = CreateButtons();
            rootStack.Children.Add(buttons);

            mainBorder.Child = rootStack;
            Content = mainBorder;
        }

        // ============================================================
        //  UI BUILDING BLOCKS
        // ============================================================

        private Border CreateTitleBar()
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleText = new TextBlock
            {
                Text = "  CONVERTER",
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
            closeBtn.Click += (s, e) => { DialogResult = false; Close(); };
            Grid.SetColumn(closeBtn, 1);
            grid.Children.Add(closeBtn);

            var border = new Border
            {
                Height = 36,
                Background = new SolidColorBrush(Color.FromRgb(22, 22, 26)),
                CornerRadius = new CornerRadius(12, 12, 0, 0),
                Child = grid
            };
            border.MouseLeftButtonDown += (s, e) => { if (e.LeftButton == MouseButtonState.Pressed) DragMove(); };

            return border;
        }

        private StackPanel CreateHeader()
        {
            var sp = new StackPanel { Margin = new Thickness(24, 16, 24, 8) };

            var title = new TextBlock
            {
                Text = "Chuyển đổi Floor / Roof / Topo",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new LinearGradientBrush(AccentBlue, AccentGreen, 0)
            };
            sp.Children.Add(title);

            var subtitle = new TextBlock
            {
                Text = "Chuyển đổi qua lại giữa Sàn, Mái và Địa hình",
                FontSize = 12,
                Foreground = new SolidColorBrush(TextSecondary),
                Margin = new Thickness(0, 4, 0, 0)
            };
            sp.Children.Add(subtitle);

            return sp;
        }

        private Border CreateSourceInfo(string sourceTypes, int count)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal };

            var icon = new TextBlock
            {
                Text = "📊",
                FontSize = 16,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            sp.Children.Add(icon);

            var info = new TextBlock
            {
                Text = $"Đã chọn: {count} phần tử  —  {sourceTypes}",
                FontSize = 13,
                Foreground = new SolidColorBrush(TextPrimary),
                VerticalAlignment = VerticalAlignment.Center
            };
            sp.Children.Add(info);

            return new Border
            {
                Background = new SolidColorBrush(BgCard),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(14, 10, 14, 10),
                Margin = new Thickness(24, 4, 24, 8),
                Child = sp
            };
        }

        private StackPanel CreateTargetSection(bool topoAvailable)
        {
            var sp = new StackPanel { Margin = new Thickness(24, 8, 24, 4) };

            var label = new TextBlock
            {
                Text = "Chuyển thành:",
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(TextPrimary),
                Margin = new Thickness(0, 0, 0, 10)
            };
            sp.Children.Add(label);

            rbFloor = CreateOptionCard("🏗️", "Floor (Sàn)", "Sàn nhấc điểm với SlabShapeEditor", "TargetGrp");
            rbRoof = CreateOptionCard("🏠", "Roof (Mái)", "Mái nhấc điểm FootPrintRoof", "TargetGrp");
            rbTopo = CreateOptionCard("🌍", "TopoSolid (Địa hình)", "Bề mặt địa hình Revit 2024+", "TargetGrp");

            rbFloor.IsChecked = true;

            if (!topoAvailable)
            {
                rbTopo.IsEnabled = false;
                rbTopo.Opacity = 0.4;
                rbTopo.ToolTip = "TopoSolid chỉ khả dụng từ Revit 2024 trở lên";
            }

            sp.Children.Add(rbFloor);
            sp.Children.Add(rbRoof);
            sp.Children.Add(rbTopo);

            return sp;
        }

        private RadioButton CreateOptionCard(string icon, string title, string description, string group)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(36) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var iconText = new TextBlock
            {
                Text = icon,
                FontSize = 18,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            Grid.SetColumn(iconText, 0);
            grid.Children.Add(iconText);

            var textStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            textStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(TextPrimary)
            });
            textStack.Children.Add(new TextBlock
            {
                Text = description,
                FontSize = 11,
                Foreground = new SolidColorBrush(TextSecondary)
            });
            Grid.SetColumn(textStack, 1);
            grid.Children.Add(textStack);

            var border = new Border
            {
                Background = new SolidColorBrush(BgCard),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 2, 0, 2),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Child = grid,
                Cursor = Cursors.Hand
            };

            var rb = new RadioButton
            {
                GroupName = group,
                Content = border,
                Margin = new Thickness(0),
                Padding = new Thickness(0),
                // Hide the default radio bullet
                Template = CreateRadioTemplate()
            };

            // Visual feedback for checked state
            rb.Checked += (s, e) =>
            {
                border.BorderBrush = new SolidColorBrush(AccentBlue);
                border.Background = new SolidColorBrush(Color.FromRgb(35, 45, 60));
            };
            rb.Unchecked += (s, e) =>
            {
                border.BorderBrush = new SolidColorBrush(BorderColor);
                border.Background = new SolidColorBrush(BgCard);
            };

            return rb;
        }

        private StackPanel CreateOptions()
        {
            var sp = new StackPanel { Margin = new Thickness(24, 8, 24, 8) };

            chkDelete = new CheckBox
            {
                Foreground = new SolidColorBrush(TextPrimary),
                FontSize = 12,
                Margin = new Thickness(4, 0, 0, 0),
                Cursor = Cursors.Hand
            };
            // Use a TextBlock for checkbox content so it inherits foreground
            var chkContent = new TextBlock
            {
                Text = "Xóa phần tử gốc sau khi chuyển đổi",
                Foreground = new SolidColorBrush(TextPrimary),
                FontSize = 12
            };
            chkDelete.Content = chkContent;

            sp.Children.Add(chkDelete);
            return sp;
        }

        private Grid CreateButtons()
        {
            var grid = new Grid { Margin = new Thickness(24, 12, 24, 20) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var btnConvert = new Button
            {
                Content = "  Chuyển đổi  ",
                Height = 36,
                MinWidth = 120,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Background = new LinearGradientBrush(AccentBlue, Color.FromRgb(36, 120, 220), 90),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 8, 0)
            };
            btnConvert.Click += BtnConvert_Click;
            Grid.SetColumn(btnConvert, 1);
            grid.Children.Add(btnConvert);

            var btnCancel = new Button
            {
                Content = "  Hủy  ",
                Height = 36,
                MinWidth = 80,
                FontSize = 13,
                Background = new SolidColorBrush(BgCard),
                Foreground = new SolidColorBrush(TextPrimary),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            Grid.SetColumn(btnCancel, 2);
            grid.Children.Add(btnCancel);

            return grid;
        }

        // ============================================================
        //  HELPERS
        // ============================================================

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (rbFloor.IsChecked == true) SelectedTargetType = ConvertTarget.Floor;
            else if (rbRoof.IsChecked == true) SelectedTargetType = ConvertTarget.Roof;
            else if (rbTopo.IsChecked == true) SelectedTargetType = ConvertTarget.TopoSolid;

            DeleteSource = chkDelete.IsChecked == true;
            DialogResult = true;
            Close();
        }

        private Separator CreateSeparator()
        {
            return new Separator
            {
                Background = new SolidColorBrush(BorderColor),
                Margin = new Thickness(24, 4, 24, 4),
                Height = 1
            };
        }

        /// <summary>
        /// Create a minimal ControlTemplate for RadioButton that hides the default bullet.
        /// </summary>
        private ControlTemplate CreateRadioTemplate()
        {
            var template = new ControlTemplate(typeof(RadioButton));
            var factory = new FrameworkElementFactory(typeof(ContentPresenter));
            factory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            template.VisualTree = factory;
            return template;
        }
    }
}
