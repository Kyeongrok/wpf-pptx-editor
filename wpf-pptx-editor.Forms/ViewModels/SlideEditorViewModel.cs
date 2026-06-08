using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using wpf_pptx_editor.Forms.Models;

namespace wpf_pptx_editor.Forms.ViewModels;

public partial class SlideEditorViewModel : ObservableObject
{
    // ── Undo 스택 ─────────────────────────────────────────────────────────────

    private readonly Stack<IUndoableAction> _undoStack = new();

    public bool CanUndo => _undoStack.Count > 0;

    public void PushUndo(IUndoableAction action)
    {
        _undoStack.Push(action);
        OnPropertyChanged(nameof(CanUndo));
    }

    [RelayCommand(CanExecute = nameof(CanUndo))]
    private void Undo()
    {
        if (_undoStack.TryPop(out var action))
        {
            action.Undo();
            OnPropertyChanged(nameof(CanUndo));
        }
    }

    // ── 상태 ──────────────────────────────────────────────────────────────────

    [ObservableProperty] private ObservableCollection<EditableShape> _shapes = new();
    [ObservableProperty] private EditableShape? _selectedShape;
    [ObservableProperty] private DrawingTool _currentTool = DrawingTool.Select;
    [ObservableProperty] private double _slideWidth = 960;
    [ObservableProperty] private double _slideHeight = 540;

    // 현재 그리기 색상
    [ObservableProperty] private Color _fillColor = Color.FromRgb(0x44, 0x72, 0xC4);
    [ObservableProperty] private Color _strokeColor = Color.FromRgb(0, 0, 0);
    [ObservableProperty] private double _strokeWidthPt = 1.0;
    [ObservableProperty] private double _fontSizePt = 12.0;
    [ObservableProperty] private string _shapeText = "";

    // Hex 입력용
    public string FillColorHex
    {
        get => $"{FillColor.R:X2}{FillColor.G:X2}{FillColor.B:X2}";
        set { if (TryParseHex(value, out var c)) FillColor = c; }
    }

    public string StrokeColorHex
    {
        get => $"{StrokeColor.R:X2}{StrokeColor.G:X2}{StrokeColor.B:X2}";
        set { if (TryParseHex(value, out var c)) StrokeColor = c; }
    }

    public Cursor CanvasCursor => CurrentTool == DrawingTool.Select
        ? Cursors.Arrow : Cursors.Cross;

    // ── 속성 변경 콜백 ────────────────────────────────────────────────────────

    partial void OnFillColorChanged(Color value)
    {
        if (SelectedShape != null)
        {
            var old = SelectedShape.FillColor;
            SelectedShape.FillColor = value;
            PushUndo(new ChangePropertyAction<Color>(
                v => SelectedShape.FillColor = v, old, "채우기 색상"));
        }
        OnPropertyChanged(nameof(FillColorHex));
    }

    partial void OnStrokeColorChanged(Color value)
    {
        if (SelectedShape != null)
        {
            var old = SelectedShape.StrokeColor;
            SelectedShape.StrokeColor = value;
            PushUndo(new ChangePropertyAction<Color>(
                v => SelectedShape.StrokeColor = v, old, "테두리 색상"));
        }
        OnPropertyChanged(nameof(StrokeColorHex));
    }

    partial void OnShapeTextChanged(string value)
    {
        if (SelectedShape != null)
        {
            var old = SelectedShape.Text;
            SelectedShape.Text = value;
            // 텍스트는 타이핑마다 기록하면 너무 많아지므로 별도 처리 가능
        }
    }

    partial void OnFontSizePtChanged(double value)
    {
        if (SelectedShape != null)
        {
            var old = SelectedShape.FontSizePt;
            SelectedShape.FontSizePt = value;
            PushUndo(new ChangePropertyAction<double>(
                v => SelectedShape.FontSizePt = v, old, "폰트 크기"));
        }
    }

    partial void OnSelectedShapeChanged(EditableShape? value)
    {
        if (value == null) return;
        _fillColor = value.FillColor;
        _strokeColor = value.StrokeColor;
        _shapeText = value.Text;
        _fontSizePt = value.FontSizePt;
        OnPropertyChanged(nameof(FillColor));
        OnPropertyChanged(nameof(StrokeColor));
        OnPropertyChanged(nameof(ShapeText));
        OnPropertyChanged(nameof(FontSizePt));
    }

    // ── 색상 스워치 적용 ─────────────────────────────────────────────────────

    public void ApplyColor(Color c)
    {
        if (SelectedShape != null)
        {
            var old = SelectedShape.FillColor;
            SelectedShape.FillColor = c;
            _fillColor = c;
            OnPropertyChanged(nameof(FillColor));
            OnPropertyChanged(nameof(FillColorHex));
            PushUndo(new ChangePropertyAction<Color>(
                v => SelectedShape.FillColor = v, old, "채우기 색상"));
        }
        else
        {
            FillColor = c;
        }
    }

    // ── 도형 추가 ────────────────────────────────────────────────────────────

    public void AddShape(double x, double y, double w, double h)
    {
        var kind = CurrentTool switch
        {
            DrawingTool.RoundedRect => ShapeKind.RoundRect,
            DrawingTool.Ellipse     => ShapeKind.Ellipse,
            DrawingTool.Arrow       => ShapeKind.Line,
            DrawingTool.TextBox     => ShapeKind.Rect,
            _                       => ShapeKind.Rect
        };

        double corner = kind == ShapeKind.RoundRect ? Math.Min(w, h) * 0.15 : 0;

        var shape = new EditableShape
        {
            X = x, Y = y, Width = w, Height = h,
            Kind = kind, CornerRadius = corner,
            FillColor = CurrentTool == DrawingTool.TextBox ? Colors.Transparent : FillColor,
            StrokeColor = StrokeColor,
            StrokeWidthPt = StrokeWidthPt,
            FontSizePt = FontSizePt
        };

        Shapes.Add(shape);
        PushUndo(new AddShapeAction(Shapes, shape));
        SelectedShape = shape;
        CurrentTool = DrawingTool.Select;
    }

    // 선 추가 (시작점/끝점 기반)
    public void AddLine(double x1, double y1, double x2, double y2)
    {
        double x = Math.Min(x1, x2), y = Math.Min(y1, y2);
        double w = Math.Abs(x2 - x1), h = Math.Abs(y2 - y1);
        if (w < 2 && h < 2) return;

        var shape = new EditableShape
        {
            X = x, Y = y, Width = Math.Max(w, 2), Height = Math.Max(h, 2),
            Kind = ShapeKind.Line,
            FlipH = x1 > x2,
            FlipV = y1 > y2,
            FillColor = Colors.Transparent,
            StrokeColor = StrokeColor,
            StrokeWidthPt = StrokeWidthPt
        };

        Shapes.Add(shape);
        PushUndo(new AddShapeAction(Shapes, shape));
        SelectedShape = shape;
        CurrentTool = DrawingTool.Select;
    }

    // 이동 완료 시 캔버스에서 호출
    public void RecordMove(EditableShape shape, double oldX, double oldY)
    {
        if (Math.Abs(shape.X - oldX) < 0.5 && Math.Abs(shape.Y - oldY) < 0.5) return;
        PushUndo(new MoveShapeAction(shape, oldX, oldY));
    }

    // ── 커맨드 ──────────────────────────────────────────────────────────────

    [RelayCommand]
    private void DeleteSelected()
    {
        if (SelectedShape is not { } sel) return;
        int idx = Shapes.IndexOf(sel);
        Shapes.Remove(sel);
        SelectedShape = null;
        PushUndo(new DeleteShapeAction(Shapes, sel, idx));
    }

    [RelayCommand]
    private void SelectTool(string toolName)
    {
        if (Enum.TryParse<DrawingTool>(toolName, out var tool))
            CurrentTool = tool;
        OnPropertyChanged(nameof(CanvasCursor));
    }

    // ── 슬라이드 변환 ────────────────────────────────────────────────────────

    public void LoadFromSlideInfo(SlideInfo slide)
    {
        SelectedShape = null;
        Shapes.Clear();
        _undoStack.Clear();
        OnPropertyChanged(nameof(CanUndo));
        SlideWidth = slide.Width;
        SlideHeight = slide.Height;
        foreach (var s in slide.Shapes)
            Shapes.Add(EditableShape.FromShapeInfo(s));
    }

    public SlideInfo ToSlideInfo()
        => new(SlideWidth, SlideHeight, Shapes.Select(s => s.ToShapeInfo()).ToList());

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────

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
