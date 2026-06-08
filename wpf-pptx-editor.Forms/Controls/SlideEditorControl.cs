using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using wpf_pptx_editor.Forms.Models;
using wpf_pptx_editor.Forms.ViewModels;

namespace wpf_pptx_editor.Forms.Controls;

public partial class SlideEditorControl : UserControl
{
    private SlideEditorViewModel? _vm;

    // 마우스 상태
    private enum MouseOp { None, Drawing, Moving }
    private MouseOp _op = MouseOp.None;
    private Point _dragStart;
    private Point _shapeOrigin;

    // 그리기 미리보기
    private Rectangle? _preview;

    // 도형 → WPF 요소 매핑
    private readonly Dictionary<Guid, (UIElement root, TextBlock? tb)> _elements = new();

    public SlideEditorControl()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (_vm != null)
        {
            _vm.Shapes.CollectionChanged -= OnShapesChanged;
            _vm.PropertyChanged -= OnVmPropertyChanged;
        }

        _vm = DataContext as SlideEditorViewModel;

        if (_vm != null)
        {
            _vm.Shapes.CollectionChanged += OnShapesChanged;
            _vm.PropertyChanged += OnVmPropertyChanged;
            RebuildCanvas();
        }
    }

    private void OnShapesChanged(object? s, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems != null)
            foreach (EditableShape shape in e.NewItems) AddShapeElement(shape);

        if (e.OldItems != null)
            foreach (EditableShape shape in e.OldItems) RemoveShapeElement(shape.Id);

        if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Reset)
            RebuildCanvas();
    }

    private void OnVmPropertyChanged(object? s, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SlideEditorViewModel.SelectedShape))
            UpdateHandles();
    }

    // ── 캔버스 재구성 ─────────────────────────────────────────────────────────

    public void RebuildCanvas()
    {
        ShapeCanvas.Children.Clear();
        _elements.Clear();

        if (_vm == null) return;
        foreach (var shape in _vm.Shapes)
            AddShapeElement(shape);

        UpdateHandles();
    }

    private void AddShapeElement(EditableShape shape)
    {
        var (root, tb) = CreateWpfElement(shape);
        Canvas.SetLeft(root, shape.X);
        Canvas.SetTop(root, shape.Y);
        ShapeCanvas.Children.Add(root);
        _elements[shape.Id] = (root, tb);

        // 속성 변경 시 동기화
        shape.PropertyChanged += (_, e) => SyncElement(shape, e.PropertyName);

        // 클릭으로 선택
        root.MouseLeftButtonDown += (s2, ev) =>
        {
            ev.Handled = true;
            if (_vm?.CurrentTool == DrawingTool.Select)
                SelectShape(shape);
        };
    }

    private void RemoveShapeElement(Guid id)
    {
        if (_elements.TryGetValue(id, out var pair))
        {
            ShapeCanvas.Children.Remove(pair.root);
            _elements.Remove(id);
        }
        UpdateHandles();
    }

    private static (UIElement root, TextBlock? tb) CreateWpfElement(EditableShape s)
    {
        UIElement shape = s.Kind switch
        {
            ShapeKind.Ellipse => new Ellipse
            {
                Width = s.Width, Height = s.Height,
                Fill = new SolidColorBrush(s.FillColor),
                Stroke = new SolidColorBrush(s.StrokeColor),
                StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0
            },
            ShapeKind.Line => new Line
            {
                X1 = s.FlipH ? s.Width : 0,
                Y1 = s.FlipV ? s.Height : 0,
                X2 = s.FlipH ? 0 : s.Width,
                Y2 = s.FlipV ? 0 : s.Height,
                Stroke = new SolidColorBrush(s.StrokeColor),
                StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0
            },
            _ => new Rectangle
            {
                Width = s.Width, Height = s.Height,
                Fill = new SolidColorBrush(s.FillColor),
                Stroke = new SolidColorBrush(s.StrokeColor),
                StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0,
                RadiusX = s.CornerRadius, RadiusY = s.CornerRadius
            }
        };

        if (s.Kind == ShapeKind.Line)
            return (shape, null);

        // 텍스트 오버레이
        var tb = new TextBlock
        {
            Text = s.Text,
            TextAlignment = s.HAlign switch
            {
                HTextAlign.Left => TextAlignment.Left,
                HTextAlign.Right => TextAlignment.Right,
                _ => TextAlignment.Center
            },
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = s.VAlign switch
            {
                VTextAlign.Top => VerticalAlignment.Top,
                VTextAlign.Bottom => VerticalAlignment.Bottom,
                _ => VerticalAlignment.Center
            },
            FontSize = s.FontSizePt * 96.0 / 72.0,
            Foreground = Brushes.Black,
            Padding = new Thickness(4),
            IsHitTestVisible = false
        };

        var grid = new Grid { Width = s.Width, Height = s.Height };
        grid.Children.Add(shape);
        grid.Children.Add(tb);

        return (grid, tb);
    }

    private void SyncElement(EditableShape s, string? prop)
    {
        if (!_elements.TryGetValue(s.Id, out var pair)) return;

        var root = pair.root;
        Canvas.SetLeft(root, s.X);
        Canvas.SetTop(root, s.Y);

        UIElement? shapeEl = root is Grid g ? g.Children[0] : root;

        if (shapeEl is Line ln)
        {
            ln.X1 = s.FlipH ? s.Width : 0;
            ln.Y1 = s.FlipV ? s.Height : 0;
            ln.X2 = s.FlipH ? 0 : s.Width;
            ln.Y2 = s.FlipV ? 0 : s.Height;
            ln.Stroke = new SolidColorBrush(s.StrokeColor);
            ln.StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0;
        }
        else if (shapeEl is Rectangle r)
        {
            r.Width = s.Width; r.Height = s.Height;
            r.Fill = new SolidColorBrush(s.FillColor);
            r.Stroke = new SolidColorBrush(s.StrokeColor);
            r.StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0;
            r.RadiusX = s.CornerRadius; r.RadiusY = s.CornerRadius;
            if (root is Grid gr) { gr.Width = s.Width; gr.Height = s.Height; }
        }
        else if (shapeEl is Ellipse el)
        {
            el.Width = s.Width; el.Height = s.Height;
            el.Fill = new SolidColorBrush(s.FillColor);
            el.Stroke = new SolidColorBrush(s.StrokeColor);
            el.StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0;
            if (root is Grid gr) { gr.Width = s.Width; gr.Height = s.Height; }
        }

        if (pair.tb != null)
        {
            pair.tb.Text = s.Text;
            pair.tb.FontSize = s.FontSizePt * 96.0 / 72.0;
            pair.tb.TextAlignment = s.HAlign switch
            {
                HTextAlign.Left => TextAlignment.Left,
                HTextAlign.Right => TextAlignment.Right,
                _ => TextAlignment.Center
            };
        }

        UpdateHandles();
    }

    // ── 선택 핸들 ─────────────────────────────────────────────────────────────

    private void UpdateHandles()
    {
        HandleCanvas.Children.Clear();
        if (_vm?.SelectedShape is not { } sel) return;

        double x = sel.X, y = sel.Y, w = sel.Width, h = sel.Height;

        // 점선 경계
        var border = new Rectangle
        {
            Width = w, Height = h,
            Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
            StrokeThickness = 1.5,
            StrokeDashArray = new DoubleCollection { 4, 2 },
            Fill = Brushes.Transparent
        };
        Canvas.SetLeft(border, x); Canvas.SetTop(border, y);
        HandleCanvas.Children.Add(border);

        // 8개 핸들
        foreach (var (hx, hy) in HandlePositions(x, y, w, h))
        {
            var handle = new Rectangle
            {
                Width = 8, Height = 8,
                Fill = Brushes.White,
                Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
                StrokeThickness = 1.5
            };
            Canvas.SetLeft(handle, hx - 4); Canvas.SetTop(handle, hy - 4);
            HandleCanvas.Children.Add(handle);
        }
    }

    private static IEnumerable<(double x, double y)> HandlePositions(
        double x, double y, double w, double h)
    {
        yield return (x,       y      ); yield return (x+w/2, y      ); yield return (x+w, y      );
        yield return (x,       y+h/2  );                                 yield return (x+w, y+h/2  );
        yield return (x,       y+h    ); yield return (x+w/2, y+h    ); yield return (x+w, y+h    );
    }

    // ── 마우스 이벤트 ─────────────────────────────────────────────────────────

    private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (_vm == null) return;
        _dragStart = e.GetPosition(ShapeCanvas);
        ShapeCanvas.CaptureMouse();

        if (_vm.CurrentTool == DrawingTool.Select)
        {
            // 빈 영역 클릭 → 선택 해제
            SelectShape(null);
            return;
        }

        // 그리기 시작
        _op = MouseOp.Drawing;
        _preview = new Rectangle
        {
            Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
            StrokeThickness = 1,
            StrokeDashArray = new DoubleCollection { 4, 2 },
            Fill = new SolidColorBrush(Color.FromArgb(40, 0x00, 0x78, 0xD4))
        };
        PreviewCanvas.Children.Add(_preview);
        UpdatePreview(_dragStart, _dragStart);
    }

    private void Canvas_MouseMove(object sender, MouseEventArgs e)
    {
        if (_vm == null || _op == MouseOp.None) return;
        var cur = e.GetPosition(ShapeCanvas);

        if (_op == MouseOp.Moving && _vm.SelectedShape is { } sel)
        {
            sel.X = _shapeOrigin.X + (cur.X - _dragStart.X);
            sel.Y = _shapeOrigin.Y + (cur.Y - _dragStart.Y);
        }
        else if (_op == MouseOp.Drawing)
        {
            UpdatePreview(_dragStart, cur);
        }
    }

    private void Canvas_MouseUp(object sender, MouseButtonEventArgs e)
    {
        ShapeCanvas.ReleaseMouseCapture();
        if (_vm == null) { _op = MouseOp.None; return; }

        var cur = e.GetPosition(ShapeCanvas);

        if (_op == MouseOp.Drawing)
        {
            if (_vm.CurrentTool == DrawingTool.Arrow)
            {
                // 선: 실제 시작/끝점 전달 (방향 보존)
                if (Math.Abs(cur.X - _dragStart.X) > 2 || Math.Abs(cur.Y - _dragStart.Y) > 2)
                    _vm.AddLine(_dragStart.X, _dragStart.Y, cur.X, cur.Y);
            }
            else
            {
                var (x, y, w, h) = NormalizeRect(_dragStart, cur);
                if (w > 4 && h > 4)
                    _vm.AddShape(x, y, w, h);
            }

            if (_preview != null)
                PreviewCanvas.Children.Remove(_preview);
            _preview = null;
        }
        else if (_op == MouseOp.Moving && _vm.SelectedShape is { } moved)
        {
            // 이동 완료 → Undo 스택에 기록
            _vm.RecordMove(moved, _shapeOrigin.X, _shapeOrigin.Y);
        }

        _op = MouseOp.None;
    }

    private void UpdatePreview(Point a, Point b)
    {
        if (_preview == null) return;
        var (x, y, w, h) = NormalizeRect(a, b);
        Canvas.SetLeft(_preview, x); Canvas.SetTop(_preview, y);
        _preview.Width = w; _preview.Height = h;
    }

    private static (double x, double y, double w, double h) NormalizeRect(Point a, Point b)
    {
        double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y);
        double w = Math.Abs(b.X - a.X), h = Math.Abs(b.Y - a.Y);
        return (x, y, w, h);
    }

    // ── 선택 ────────────────────────────────────────────────────────────────

    private void SelectShape(EditableShape? shape)
    {
        if (_vm == null) return;

        if (_vm.SelectedShape != null)
            _vm.SelectedShape.IsSelected = false;

        _vm.SelectedShape = shape;

        if (shape != null)
        {
            shape.IsSelected = true;
            _op = MouseOp.Moving;
            _shapeOrigin = new Point(shape.X, shape.Y);
        }

        UpdateHandles();
    }
}
