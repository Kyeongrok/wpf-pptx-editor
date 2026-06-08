using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using wpf_pptx_editor.Forms.Models;

namespace wpf_pptx_editor.Forms.Services;

public class SlideRenderer
{
    private const double PtToDip = 96.0 / 72.0;

    public BitmapImage RenderToBitmap(SlideInfo slide, int targetWidth)
    {
        double scale = targetWidth / slide.Width;
        int targetHeight = (int)Math.Round(slide.Height * scale);

        var canvas = BuildCanvas(slide, scale);
        canvas.Measure(new Size(canvas.Width, canvas.Height));
        canvas.Arrange(new Rect(0, 0, canvas.Width, canvas.Height));

        var rtb = new RenderTargetBitmap(targetWidth, targetHeight, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(canvas);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(rtb));
        using var ms = new MemoryStream();
        encoder.Save(ms);
        ms.Position = 0;

        var bmp = new BitmapImage();
        bmp.BeginInit();
        bmp.StreamSource = ms;
        bmp.CacheOption = BitmapCacheOption.OnLoad;
        bmp.EndInit();
        bmp.Freeze();
        return bmp;
    }

    private static Canvas BuildCanvas(SlideInfo slide, double scale)
    {
        var canvas = new Canvas
        {
            Width = slide.Width * scale,
            Height = slide.Height * scale,
            Background = Brushes.White,
            ClipToBounds = true
        };

        foreach (var shape in slide.Shapes)
            AddShape(canvas, shape, scale);

        return canvas;
    }

    private static void AddShape(Canvas canvas, ShapeInfo shape, double scale)
    {
        UIElement? el = shape.Kind switch
        {
            ShapeKind.Ellipse => MakeEllipse(shape, scale),
            ShapeKind.Line => MakeLine(shape, scale),
            _ => MakeRect(shape, scale)
        };

        if (el != null) canvas.Children.Add(el);

        if (shape.Text != null && shape.Kind != ShapeKind.Line)
        {
            var textEl = MakeTextOverlay(shape, scale);
            if (textEl != null) canvas.Children.Add(textEl);
        }
    }

    private static Rectangle MakeRect(ShapeInfo s, double scale)
    {
        var r = new Rectangle
        {
            Width = s.Width * scale,
            Height = s.Height * scale,
            Fill = Brush(s.FillColor),
            Stroke = Brush(s.StrokeColor),
            StrokeThickness = s.StrokeWidthPt * PtToDip * scale,
            RadiusX = s.CornerRadius * scale,
            RadiusY = s.CornerRadius * scale
        };
        Place(r, s.X * scale, s.Y * scale);
        return r;
    }

    private static Ellipse MakeEllipse(ShapeInfo s, double scale)
    {
        var e = new Ellipse
        {
            Width = s.Width * scale,
            Height = s.Height * scale,
            Fill = Brush(s.FillColor),
            Stroke = Brush(s.StrokeColor),
            StrokeThickness = s.StrokeWidthPt * PtToDip * scale
        };
        Place(e, s.X * scale, s.Y * scale);
        return e;
    }

    private static Line MakeLine(ShapeInfo s, double scale)
    {
        double x1 = s.FlipH ? (s.X + s.Width) * scale : s.X * scale;
        double y1 = s.FlipV ? (s.Y + s.Height) * scale : s.Y * scale;
        double x2 = s.FlipH ? s.X * scale : (s.X + s.Width) * scale;
        double y2 = s.FlipV ? s.Y * scale : (s.Y + s.Height) * scale;

        return new Line
        {
            X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
            Stroke = Brush(s.StrokeColor ?? System.Windows.Media.Color.FromRgb(0, 0, 0)),
            StrokeThickness = s.StrokeWidthPt * PtToDip * scale
        };
    }

    private static Border? MakeTextOverlay(ShapeInfo s, double scale)
    {
        if (s.Text == null || s.Text.Paragraphs.Count == 0) return null;

        var vAlign = s.Text.VAlign switch
        {
            VTextAlign.Top => VerticalAlignment.Top,
            VTextAlign.Bottom => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };

        var stack = new StackPanel { VerticalAlignment = vAlign };

        foreach (var para in s.Text.Paragraphs)
        {
            var hAlign = para.HAlign switch
            {
                HTextAlign.Left => TextAlignment.Left,
                HTextAlign.Right => TextAlignment.Right,
                HTextAlign.Justify => TextAlignment.Justify,
                _ => TextAlignment.Center
            };

            var tb = new TextBlock
            {
                TextAlignment = hAlign,
                TextWrapping = TextWrapping.Wrap
            };

            foreach (var run in para.Runs)
            {
                tb.Inlines.Add(new System.Windows.Documents.Run(run.Text)
                {
                    FontSize = run.FontSizePt * PtToDip * scale,
                    FontWeight = run.Bold ? FontWeights.Bold : FontWeights.Normal,
                    FontStyle = run.Italic ? FontStyles.Italic : FontStyles.Normal,
                    Foreground = run.FontColor.HasValue
                        ? new SolidColorBrush(run.FontColor.Value)
                        : Brushes.Black
                });
            }

            stack.Children.Add(tb);
        }

        var border = new Border
        {
            Width = s.Width * scale,
            Height = s.Height * scale,
            Padding = new Thickness(4 * scale),
            Child = stack
        };
        Place(border, s.X * scale, s.Y * scale);
        return border;
    }

    private static void Place(UIElement el, double left, double top)
    {
        Canvas.SetLeft(el, left);
        Canvas.SetTop(el, top);
    }

    private static Brush Brush(System.Windows.Media.Color? c)
        => c.HasValue ? new SolidColorBrush(c.Value) : Brushes.Transparent;
}
