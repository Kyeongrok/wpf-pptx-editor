using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;

namespace wpf_pptx_editor.Forms.Models;

public enum DrawingTool { Select, Rectangle, RoundedRect, Ellipse, Arrow, TextBox }

public partial class EditableShape : ObservableObject
{
    public Guid Id { get; } = Guid.NewGuid();

    [ObservableProperty] private double _x;
    [ObservableProperty] private double _y;
    [ObservableProperty] private double _width;
    [ObservableProperty] private double _height;
    [ObservableProperty] private ShapeKind _kind = ShapeKind.Rect;
    [ObservableProperty] private double _cornerRadius;
    [ObservableProperty] private Color _fillColor = Color.FromRgb(0x44, 0x72, 0xC4);
    [ObservableProperty] private Color _strokeColor = Color.FromRgb(0, 0, 0);
    [ObservableProperty] private double _strokeWidthPt = 1.0;
    [ObservableProperty] private string _text = "";
    [ObservableProperty] private HTextAlign _hAlign = HTextAlign.Center;
    [ObservableProperty] private VTextAlign _vAlign = VTextAlign.Center;
    [ObservableProperty] private double _fontSizePt = 12.0;
    [ObservableProperty] private bool _isSelected;

    // 선 방향 (Line 전용)
    [ObservableProperty] private bool _flipH; // 오른쪽→왼쪽
    [ObservableProperty] private bool _flipV; // 아래→위

    public static EditableShape FromShapeInfo(ShapeInfo s) => new()
    {
        X = s.X, Y = s.Y, Width = s.Width, Height = s.Height,
        Kind = s.Kind, CornerRadius = s.CornerRadius,
        FlipH = s.FlipH, FlipV = s.FlipV,
        FillColor = s.FillColor ?? Color.FromRgb(0xD0, 0xD0, 0xD0),
        StrokeColor = s.StrokeColor ?? Color.FromRgb(0, 0, 0),
        StrokeWidthPt = s.StrokeWidthPt,
        Text = s.Text?.Paragraphs.FirstOrDefault()?.Runs.FirstOrDefault()?.Text ?? "",
        HAlign = s.Text?.Paragraphs.FirstOrDefault()?.HAlign ?? HTextAlign.Center,
        VAlign = s.Text?.VAlign ?? VTextAlign.Center,
        FontSizePt = s.Text?.Paragraphs.FirstOrDefault()?.Runs.FirstOrDefault()?.FontSizePt ?? 12.0
    };

    public ShapeInfo ToShapeInfo()
    {
        TextInfo? textInfo = null;
        if (!string.IsNullOrEmpty(Text))
        {
            var run = new TextRunInfo(Text, FontSizePt, false, false, Color.FromRgb(0, 0, 0));
            textInfo = new TextInfo(new[] { new TextParagraph(new[] { run }, HAlign) }, VAlign);
        }
        return new ShapeInfo(X, Y, Width, Height, Kind, CornerRadius,
            FillColor, StrokeColor, StrokeWidthPt, textInfo, FlipH, FlipV);
    }
}
