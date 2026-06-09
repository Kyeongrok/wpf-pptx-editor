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
    private enum MouseOp { None, Drawing, Moving, ResizingLine }
    private MouseOp _op = MouseOp.None;
    private Point _dragStart;
    private Point _shapeOrigin;
    private bool _resizingLineStart;   // true=시작점, false=끝점 드래그 중
    private EditableShape? _resizingSnapShape;

    // 그리기 미리보기
    private UIElement? _preview;

    // 연결선 스냅
    private EditableShape? _snapStartShape;
    private EditableShape? _snapEndShape;
    private Point _lineStart;          // Arrow 그리기 실제 시작점 (스냅 보정 포함)
    private double _snapStartAnchorX = 0.5, _snapStartAnchorY = 0.5;
    private double _snapEndAnchorX   = 0.5, _snapEndAnchorY   = 0.5;
    private double _resizingSnapAnchorX = 0.5, _resizingSnapAnchorY = 0.5;
    private System.Windows.Shapes.Ellipse? _snapIndicator;

    // 텍스트 인라인 편집
    private TextBox? _editBox;
    private EditableShape? _editingShape;
    public bool IsTextEditing => _editBox != null;

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

        // 클릭으로 선택 (Line은 끝점 근처면 리사이즈 모드)
        root.MouseLeftButtonDown += (s2, ev) =>
        {
            ev.Handled = true;
            if (_vm?.CurrentTool != DrawingTool.Select) return;

            _dragStart = ev.GetPosition(ShapeCanvas);

            if (shape.Kind == ShapeKind.Line)
            {
                double x1 = shape.FlipH ? shape.X + shape.Width  : shape.X;
                double y1 = shape.FlipV ? shape.Y + shape.Height : shape.Y;
                double x2 = shape.FlipH ? shape.X                : shape.X + shape.Width;
                double y2 = shape.FlipV ? shape.Y                : shape.Y + shape.Height;

                if (Distance(_dragStart, new Point(x1, y1)) <= 10)
                {
                    SelectShape(shape);
                    _op = MouseOp.ResizingLine;
                    _resizingLineStart = true;
                    ShapeCanvas.CaptureMouse();
                    return;
                }
                if (Distance(_dragStart, new Point(x2, y2)) <= 10)
                {
                    SelectShape(shape);
                    _op = MouseOp.ResizingLine;
                    _resizingLineStart = false;
                    ShapeCanvas.CaptureMouse();
                    return;
                }
            }

            SelectShape(shape);
            ShapeCanvas.CaptureMouse();
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
            ShapeKind.Line => CreateLineElement(s),
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

        // Line은 Canvas > [Line(visual), Line(hit)] 구조
        if (root is Canvas lineCanvas && lineCanvas.Children.Count == 2
            && lineCanvas.Children[0] is Line visual && lineCanvas.Children[1] is Line hitLine)
        {
            double x1 = s.FlipH ? s.Width : 0, y1 = s.FlipV ? s.Height : 0;
            double x2 = s.FlipH ? 0 : s.Width, y2 = s.FlipV ? 0 : s.Height;
            visual.X1 = hitLine.X1 = x1; visual.Y1 = hitLine.Y1 = y1;
            visual.X2 = hitLine.X2 = x2; visual.Y2 = hitLine.Y2 = y2;
            visual.Stroke = new SolidColorBrush(s.StrokeColor);
            visual.StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0;
            lineCanvas.Width = s.Width; lineCanvas.Height = s.Height;
            UpdateHandles();
            return;
        }

        UIElement? shapeEl = root is Grid g ? g.Children[0] : root;

        if (shapeEl is Rectangle r)
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

        if (sel.Kind == ShapeKind.Line)
        {
            // 선: 끝점 2개 핸들만 표시
            double x1 = sel.FlipH ? x + w : x, y1 = sel.FlipV ? y + h : y;
            double x2 = sel.FlipH ? x : x + w, y2 = sel.FlipV ? y : y + h;

            AddHandle(x1, y1);
            AddHandle(x2, y2);
        }
        else
        {
            // 일반 도형: 점선 경계 + 8개 핸들
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

            foreach (var (hx, hy) in HandlePositions(x, y, w, h))
                AddHandle(hx, hy);
        }
    }

    private void AddHandle(double cx, double cy)
    {
        var handle = new Rectangle
        {
            Width = 8, Height = 8,
            Fill = Brushes.White,
            Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
            StrokeThickness = 1.5
        };
        Canvas.SetLeft(handle, cx - 4); Canvas.SetTop(handle, cy - 4);
        HandleCanvas.Children.Add(handle);
    }

    private static IEnumerable<(double x, double y)> HandlePositions(
        double x, double y, double w, double h)
    {
        yield return (x,       y      ); yield return (x+w/2, y      ); yield return (x+w, y      );
        yield return (x,       y+h/2  );                                 yield return (x+w, y+h/2  );
        yield return (x,       y+h    ); yield return (x+w/2, y+h    ); yield return (x+w, y+h    );
    }

    // ── 텍스트 인라인 편집 ────────────────────────────────────────────────────

    public void EnterTextEditMode()
    {
        if (_vm?.SelectedShape is not { } sel) return;
        if (sel.Kind == ShapeKind.Line) return;
        if (_editBox != null) return;

        _editingShape = sel;

        _editBox = new TextBox
        {
            Width = sel.Width,
            Height = sel.Height,
            Text = sel.Text,
            TextAlignment = sel.HAlign switch
            {
                HTextAlign.Left  => TextAlignment.Left,
                HTextAlign.Right => TextAlignment.Right,
                _                => TextAlignment.Center
            },
            VerticalContentAlignment = sel.VAlign switch
            {
                VTextAlign.Top    => VerticalAlignment.Top,
                VTextAlign.Bottom => VerticalAlignment.Bottom,
                _                 => VerticalAlignment.Center
            },
            FontSize = sel.FontSizePt * 96.0 / 72.0,
            Background = new SolidColorBrush(Color.FromArgb(200, 255, 255, 255)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
            BorderThickness = new Thickness(1.5),
            Padding = new Thickness(4),
            TextWrapping = TextWrapping.Wrap,
            AcceptsReturn = false,
        };

        Canvas.SetLeft(_editBox, sel.X);
        Canvas.SetTop(_editBox, sel.Y);

        _editBox.PreviewKeyDown += OnEditBoxKeyDown;
        _editBox.LostFocus += OnEditBoxLostFocus;

        TextEditCanvas.IsHitTestVisible = true;
        TextEditCanvas.Children.Add(_editBox);
        _editBox.Focus();
        _editBox.SelectAll();
    }

    private void OnEditBoxKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)      { ExitTextEditMode(commit: false); e.Handled = true; }
        else if (e.Key == Key.Enter)  { ExitTextEditMode(commit: true);  e.Handled = true; }
    }

    private void OnEditBoxLostFocus(object sender, RoutedEventArgs e)
        => ExitTextEditMode(commit: true);

    private void ExitTextEditMode(bool commit)
    {
        if (_editBox == null) return;

        if (commit && _editingShape != null)
            _editingShape.Text = _editBox.Text;

        _editBox.PreviewKeyDown -= OnEditBoxKeyDown;
        _editBox.LostFocus      -= OnEditBoxLostFocus;
        TextEditCanvas.Children.Remove(_editBox);
        TextEditCanvas.IsHitTestVisible = false;
        _editBox = null;
        _editingShape = null;
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
        if (_vm.CurrentTool == DrawingTool.Arrow)
        {
            _snapStartShape = FindConnectableShapeAt(_dragStart);
            if (_snapStartShape != null)
            {
                var (sp, ax, ay) = NearestConnectionPoint(_snapStartShape, _dragStart);
                _lineStart = sp; _snapStartAnchorX = ax; _snapStartAnchorY = ay;
            }
            else { _lineStart = _dragStart; }

            _preview = new Line
            {
                X1 = _lineStart.X, Y1 = _lineStart.Y,
                X2 = _lineStart.X, Y2 = _lineStart.Y,
                Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),
                StrokeThickness = 1.5,
                StrokeDashArray = new DoubleCollection { 4, 2 }
            };
            PreviewCanvas.Children.Add(_preview);
        }
        else
        {
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
    }

    private void Canvas_MouseMove(object sender, MouseEventArgs e)
    {
        if (_vm == null || _op == MouseOp.None) return;
        var cur = e.GetPosition(ShapeCanvas);

        if (_op == MouseOp.ResizingLine && _vm.SelectedShape is { Kind: ShapeKind.Line } rline)
        {
            // 고정 끝점 계산
            double fx = _resizingLineStart
                ? (rline.FlipH ? rline.X : rline.X + rline.Width)
                : (rline.FlipH ? rline.X + rline.Width : rline.X);
            double fy = _resizingLineStart
                ? (rline.FlipV ? rline.Y : rline.Y + rline.Height)
                : (rline.FlipV ? rline.Y + rline.Height : rline.Y);

            _resizingSnapShape = FindConnectableShapeAt(cur);
            Point movPt;
            if (_resizingSnapShape != null)
            {
                var (rp, ax, ay) = NearestConnectionPoint(_resizingSnapShape, cur);
                movPt = rp; _resizingSnapAnchorX = ax; _resizingSnapAnchorY = ay;
                ShowSnapIndicator(movPt);
            }
            else { movPt = cur; HideSnapIndicator(); }

            if (_resizingLineStart) UpdateLineEndpoints(rline, movPt.X, movPt.Y, fx, fy);
            else                    UpdateLineEndpoints(rline, fx, fy, movPt.X, movPt.Y);
        }
        else if (_op == MouseOp.Moving && _vm.SelectedShape is { } sel)
        {
            sel.X = _shapeOrigin.X + (cur.X - _dragStart.X);
            sel.Y = _shapeOrigin.Y + (cur.Y - _dragStart.Y);
            if (sel.Kind != ShapeKind.Line)
                _vm.UpdateConnectedLines(sel);
        }
        else if (_op == MouseOp.Drawing)
        {
            if (_vm.CurrentTool == DrawingTool.Arrow)
            {
                _snapEndShape = FindConnectableShapeAt(cur, _snapStartShape);
                Point end;
                if (_snapEndShape != null)
                {
                    var (ep, ax, ay) = NearestConnectionPoint(_snapEndShape, cur);
                    end = ep; _snapEndAnchorX = ax; _snapEndAnchorY = ay;
                    ShowSnapIndicator(end);
                }
                else
                {
                    end = (Keyboard.Modifiers & ModifierKeys.Shift) != 0
                        ? SnapTo15Deg(_lineStart, cur) : cur;
                    HideSnapIndicator();
                }
                UpdatePreview(_lineStart, end);
            }
            else
            {
                UpdatePreview(_dragStart, cur);
            }
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
                var endShape = _snapEndShape;
                Point end;
                if (endShape != null)
                {
                    var (ep, ax, ay) = NearestConnectionPoint(endShape, cur);
                    end = ep; _snapEndAnchorX = ax; _snapEndAnchorY = ay;
                }
                else
                {
                    end = (Keyboard.Modifiers & ModifierKeys.Shift) != 0
                        ? SnapTo15Deg(_lineStart, cur) : cur;
                }
                if (Math.Abs(end.X - _lineStart.X) > 2 || Math.Abs(end.Y - _lineStart.Y) > 2)
                    _vm.AddLine(_lineStart.X, _lineStart.Y, end.X, end.Y,
                                _snapStartShape?.Id, endShape?.Id,
                                _snapStartAnchorX, _snapStartAnchorY,
                                _snapEndAnchorX, _snapEndAnchorY);
                HideSnapIndicator();
                _snapStartShape = null;
                _snapEndShape = null;
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
        else if (_op == MouseOp.ResizingLine && _vm.SelectedShape is { Kind: ShapeKind.Line } resized)
        {
            // 리사이즈 완료 → 연결 ID + 앵커 갱신
            if (_resizingLineStart)
            {
                resized.StartConnectedShapeId = _resizingSnapShape?.Id;
                resized.StartConnectedAnchorX = _resizingSnapAnchorX;
                resized.StartConnectedAnchorY = _resizingSnapAnchorY;
            }
            else
            {
                resized.EndConnectedShapeId = _resizingSnapShape?.Id;
                resized.EndConnectedAnchorX = _resizingSnapAnchorX;
                resized.EndConnectedAnchorY = _resizingSnapAnchorY;
            }
            HideSnapIndicator();
            _resizingSnapShape = null;
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
        if (_preview is Line ln)
        {
            ln.X1 = a.X; ln.Y1 = a.Y;
            ln.X2 = b.X; ln.Y2 = b.Y;
        }
        else if (_preview is Rectangle rect)
        {
            var (x, y, w, h) = NormalizeRect(a, b);
            Canvas.SetLeft(rect, x); Canvas.SetTop(rect, y);
            rect.Width = w; rect.Height = h;
        }
    }

    private static Point SnapTo15Deg(Point from, Point to)
    {
        double dx = to.X - from.X;
        double dy = to.Y - from.Y;
        double len = Math.Sqrt(dx * dx + dy * dy);
        if (len < 1) return to;

        double angleDeg = Math.Atan2(dy, dx) * 180.0 / Math.PI;
        double snapped = Math.Round(angleDeg / 15.0) * 15.0;
        double rad = snapped * Math.PI / 180.0;
        return new Point(from.X + len * Math.Cos(rad), from.Y + len * Math.Sin(rad));
    }

    private static Canvas CreateLineElement(EditableShape s)
    {
        double x1 = s.FlipH ? s.Width : 0, y1 = s.FlipV ? s.Height : 0;
        double x2 = s.FlipH ? 0 : s.Width, y2 = s.FlipV ? 0 : s.Height;

        var visual = new Line
        {
            X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
            Stroke = new SolidColorBrush(s.StrokeColor),
            StrokeThickness = s.StrokeWidthPt * 96.0 / 72.0,
            IsHitTestVisible = false
        };
        // 투명 히트 영역 — 실제 선 주변 10px 반경 클릭 가능
        var hit = new Line
        {
            X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
            Stroke = Brushes.Transparent,
            StrokeThickness = 10
        };

        var canvas = new Canvas { Width = s.Width, Height = s.Height };
        canvas.Children.Add(visual);
        canvas.Children.Add(hit);
        return canvas;
    }

    // ── 연결선 스냅 헬퍼 ─────────────────────────────────────────────────────

    private EditableShape? FindConnectableShapeAt(Point pt, EditableShape? exclude = null)
    {
        const double margin = 8;
        if (_vm == null) return null;
        return _vm.Shapes
            .Where(s => s.Kind != ShapeKind.Line && s != exclude)
            .LastOrDefault(s => pt.X >= s.X - margin && pt.X <= s.X + s.Width  + margin
                             && pt.Y >= s.Y - margin && pt.Y <= s.Y + s.Height + margin);
    }

    // 도형의 8개 핸들 중 마우스에 가장 가까운 점과 그 상대 앵커 반환
    private static (Point pt, double anchorX, double anchorY) NearestConnectionPoint(
        EditableShape s, Point mouse)
    {
        (double rx, double ry)[] anchors =
        {
            (0,   0  ), (0.5, 0  ), (1,   0  ),
            (0,   0.5),             (1,   0.5),
            (0,   1  ), (0.5, 1  ), (1,   1  )
        };
        var best = anchors.MinBy(a => Distance(mouse,
            new Point(s.X + a.rx * s.Width, s.Y + a.ry * s.Height)));
        var pt = new Point(s.X + best.rx * s.Width, s.Y + best.ry * s.Height);
        return (pt, best.rx, best.ry);
    }

    private void ShowSnapIndicator(Point pt)
    {
        if (_snapIndicator == null)
        {
            _snapIndicator = new System.Windows.Shapes.Ellipse
            {
                Width = 14, Height = 14,
                Fill = new SolidColorBrush(Color.FromArgb(60, 0, 178, 100)),
                Stroke = new SolidColorBrush(Color.FromRgb(0, 178, 100)),
                StrokeThickness = 2,
                IsHitTestVisible = false
            };
            PreviewCanvas.Children.Add(_snapIndicator);
        }
        Canvas.SetLeft(_snapIndicator, pt.X - 7);
        Canvas.SetTop(_snapIndicator, pt.Y - 7);
    }

    private void HideSnapIndicator()
    {
        if (_snapIndicator == null) return;
        PreviewCanvas.Children.Remove(_snapIndicator);
        _snapIndicator = null;
    }

    private static (double x, double y, double w, double h) NormalizeRect(Point a, Point b)
    {
        double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y);
        double w = Math.Abs(b.X - a.X), h = Math.Abs(b.Y - a.Y);
        return (x, y, w, h);
    }

    private static void UpdateLineEndpoints(EditableShape line, double x1, double y1, double x2, double y2)
    {
        line.X = Math.Min(x1, x2);
        line.Y = Math.Min(y1, y2);
        line.Width  = Math.Max(Math.Abs(x2 - x1), 2);
        line.Height = Math.Max(Math.Abs(y2 - y1), 2);
        line.FlipH  = x1 > x2;
        line.FlipV  = y1 > y2;
    }

    private static double Distance(Point a, Point b)
    {
        double dx = a.X - b.X, dy = a.Y - b.Y;
        return Math.Sqrt(dx * dx + dy * dy);
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
