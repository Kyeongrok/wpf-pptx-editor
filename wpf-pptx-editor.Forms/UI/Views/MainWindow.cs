using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using wpf_pptx_editor.Forms.Controls;
using wpf_pptx_editor.Forms.ViewModels;
using wpf_pptx_editor.Support.UI.Units;

namespace wpf_pptx_editor.Forms.UI.Views;

public class MainWindow : wpf_pptx_editorWindow
{
    static MainWindow()
    {
        DefaultStyleKeyProperty.OverrideMetadata(typeof(MainWindow),
            new FrameworkPropertyMetadata(typeof(MainWindow)));
    }

    public MainWindow(MainWindowViewModel viewModel)
    {
        DataContext = viewModel;
        KeyDown += (_, e) =>
        {
            if (e.Key == System.Windows.Input.Key.Delete)
                viewModel.Editor.DeleteSelectedCommand.Execute(null);

            if (e.Key == System.Windows.Input.Key.Z
                && (System.Windows.Input.Keyboard.Modifiers
                    & System.Windows.Input.ModifierKeys.Control) != 0)
            {
                viewModel.Editor.UndoCommand.Execute(null);
                e.Handled = true;
            }
        };
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();

        Wire("PART_MinimizeButton", () => WindowState = WindowState.Minimized);
        Wire("PART_MaximizeButton", () =>
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal : WindowState.Maximized);
        Wire("PART_CloseButton", Close);

        // 색상 스워치
        var vm = (DataContext as MainWindowViewModel)?.Editor;
        if (vm != null)
        {
            var palette = new[]
            {
                "#4472C4","#ED7D31","#A9D18E","#FFC000","#5B9BD5","#70AD47",
                "#FF0000","#F5A0A0","#FFFFFF","#000000","#808080","transparent"
            };
            for (int i = 1; i <= 12; i++)
            {
                int idx = i - 1;
                if (GetTemplateChild($"PART_Color{i}") is Button btn)
                {
                    btn.Click += (_, _) =>
                    {
                        if (palette[idx] == "transparent")
                            vm.ApplyColor(Colors.Transparent);
                        else if (TryParseHex(palette[idx], out var c))
                            vm.ApplyColor(c);
                    };
                }
            }
        }

        // 슬라이드 선택 시 캔버스 재구성
        if (DataContext is MainWindowViewModel mainVm)
        {
            mainVm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(MainWindowViewModel.Editor))
                    RebuildEditorCanvas();
            };
            mainVm.Editor.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(SlideEditorViewModel.Shapes))
                    RebuildEditorCanvas();
            };
        }
    }

    private void RebuildEditorCanvas()
    {
        // SlideEditorControl은 DataContext 변경 시 자동으로 RebuildCanvas를 호출하므로
        // 명시적 호출이 필요한 경우만 여기서 처리
    }

    private void Wire(string partName, Action action)
    {
        if (GetTemplateChild(partName) is Button btn)
            btn.Click += (_, _) => action();
    }

    private static bool TryParseHex(string hex, out Color c)
    {
        c = Colors.Black;
        hex = hex.TrimStart('#');
        if (hex.Length != 6) return false;
        try
        {
            c = Color.FromRgb(
                Convert.ToByte(hex[..2], 16),
                Convert.ToByte(hex[2..4], 16),
                Convert.ToByte(hex[4..6], 16));
            return true;
        }
        catch { return false; }
    }
}
