using System.Windows;

namespace wpf_pptx_editor.Support.UI.Units;

public class wpf_pptx_editorWindow : Window
{
    static wpf_pptx_editorWindow()
    {
        DefaultStyleKeyProperty.OverrideMetadata(typeof(wpf_pptx_editorWindow),
            new FrameworkPropertyMetadata(typeof(wpf_pptx_editorWindow)));
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);
        if (WindowState == WindowState.Maximized)
            MaxHeight = SystemParameters.WorkArea.Height;
        else
            MaxHeight = double.PositiveInfinity;
    }
}
