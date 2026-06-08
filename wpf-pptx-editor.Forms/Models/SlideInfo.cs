using System.Windows.Media;

namespace wpf_pptx_editor.Forms.Models;

public enum ShapeKind { Rect, RoundRect, Ellipse, Triangle, Diamond, Line, Other }
public enum HTextAlign { Left, Center, Right, Justify }
public enum VTextAlign { Top, Center, Bottom }

public record TextRunInfo(string Text, double FontSizePt, bool Bold, bool Italic, Color? FontColor);
public record TextParagraph(IReadOnlyList<TextRunInfo> Runs, HTextAlign HAlign);
public record TextInfo(IReadOnlyList<TextParagraph> Paragraphs, VTextAlign VAlign);

public record ShapeInfo(
    double X, double Y,
    double Width, double Height,
    ShapeKind Kind,
    double CornerRadius,
    Color? FillColor,
    Color? StrokeColor,
    double StrokeWidthPt,
    TextInfo? Text,
    bool FlipH = false,
    bool FlipV = false);

public record SlideInfo(
    double Width,
    double Height,
    IReadOnlyList<ShapeInfo> Shapes);
