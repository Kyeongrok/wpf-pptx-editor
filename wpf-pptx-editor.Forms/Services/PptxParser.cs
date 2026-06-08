using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using wpf_pptx_editor.Forms.Models;
using WColor = System.Windows.Media.Color;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace wpf_pptx_editor.Forms.Services;

public class PptxParser
{
    private const double EmuToDip = 96.0 / 914400.0;
    private const double EmuToPt  = 1.0 / 12700.0;

    private Dictionary<string, WColor> _themeColors = new();

    public void LoadThemeColors(PresentationDocument doc)
    {
        _themeColors = new();
        var scheme = doc.PresentationPart?.ThemePart?.Theme
            .ThemeElements?.ColorScheme;
        if (scheme == null) return;

        void Add(string key, OpenXmlElement? el)
        {
            if (el == null) return;
            var rgb = el.GetFirstChild<A.RgbColorModelHex>();
            if (rgb?.Val?.Value is string h) { _themeColors[key] = ParseHex(h); return; }
            var sys = el.GetFirstChild<A.SystemColor>();
            if (sys?.LastColor?.Value is string lc) _themeColors[key] = ParseHex(lc);
        }

        Add("dk1",     scheme.Dark1Color);
        Add("lt1",     scheme.Light1Color);
        Add("dk2",     scheme.Dark2Color);
        Add("lt2",     scheme.Light2Color);
        Add("accent1", scheme.Accent1Color);
        Add("accent2", scheme.Accent2Color);
        Add("accent3", scheme.Accent3Color);
        Add("accent4", scheme.Accent4Color);
        Add("accent5", scheme.Accent5Color);
        Add("accent6", scheme.Accent6Color);
        Add("hlink",   scheme.Hyperlink);
        Add("folHlink",scheme.FollowedHyperlinkColor);
    }

    public int GetSlideCount(PresentationDocument doc)
        => doc.PresentationPart?.Presentation.SlideIdList?.Count() ?? 0;

    public SlideInfo ParseSlide(PresentationDocument doc, int slideIndex)
    {
        var presPart = doc.PresentationPart!;
        var slideId  = presPart.Presentation.SlideIdList!
            .Elements<P.SlideId>().ElementAt(slideIndex);
        var slidePart = (SlidePart)presPart.GetPartById(slideId.RelationshipId!.Value!);

        var sz = presPart.Presentation.SlideSize!;
        double w = sz.Cx!.Value * EmuToDip;
        double h = sz.Cy!.Value * EmuToDip;

        var shapes = new List<ShapeInfo>();
        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree != null)
        {
            foreach (var child in spTree.ChildElements)
            {
                ShapeInfo? info = child switch
                {
                    P.Shape sp           => ParseShape(sp),
                    P.ConnectionShape cs => ParseConnector(cs),
                    _                    => null
                };
                if (info != null) shapes.Add(info);
            }
        }

        return new SlideInfo(w, h, shapes);
    }

    // ─── Shape ───────────────────────────────────────────────────────────────

    private ShapeInfo? ParseShape(P.Shape sp)
    {
        var spPr = sp.ShapeProperties;
        if (spPr == null) return null;

        var xfrm = spPr.GetFirstChild<A.Transform2D>();
        if (xfrm == null) return null;

        double x  = (xfrm.Offset?.X?.Value  ?? 0) * EmuToDip;
        double y  = (xfrm.Offset?.Y?.Value  ?? 0) * EmuToDip;
        double sw = (xfrm.Extents?.Cx?.Value ?? 0) * EmuToDip;
        double sh = (xfrm.Extents?.Cy?.Value ?? 0) * EmuToDip;

        var prstGeom = spPr.GetFirstChild<A.PresetGeometry>();
        var kind     = ToShapeKind(prstGeom?.Preset?.Value);
        double cr    = CalcCornerRadius(kind, prstGeom, sw, sh);

        var fill               = ParseFill(spPr);
        var (stroke, strokeW)  = ParseStroke(spPr.GetFirstChild<A.Outline>());
        var text               = sp.TextBody is { } tb ? ParseTextBody(tb) : null;

        return new ShapeInfo(x, y, sw, sh, kind, cr, fill, stroke, strokeW, text);
    }

    private ShapeInfo? ParseConnector(P.ConnectionShape cs)
    {
        var spPr = cs.ShapeProperties;
        if (spPr == null) return null;
        var xfrm = spPr.GetFirstChild<A.Transform2D>();
        if (xfrm == null) return null;

        double x  = (xfrm.Offset?.X?.Value  ?? 0) * EmuToDip;
        double y  = (xfrm.Offset?.Y?.Value  ?? 0) * EmuToDip;
        double sw = (xfrm.Extents?.Cx?.Value ?? 0) * EmuToDip;
        double sh = (xfrm.Extents?.Cy?.Value ?? 0) * EmuToDip;

        var (stroke, strokeW) = ParseStroke(spPr.GetFirstChild<A.Outline>());
        return new ShapeInfo(x, y, sw, sh, ShapeKind.Line, 0,
            null, stroke ?? WColor.FromRgb(0, 0, 0), Math.Max(strokeW, 1.0), null);
    }

    // ─── Geometry ─────────────────────────────────────────────────────────────

    private static ShapeKind ToShapeKind(A.ShapeTypeValues? v)
    {
        if (v == null) return ShapeKind.Other;
        if (v.Equals(A.ShapeTypeValues.Rectangle))      return ShapeKind.Rect;
        if (v.Equals(A.ShapeTypeValues.RoundRectangle)) return ShapeKind.RoundRect;
        if (v.Equals(A.ShapeTypeValues.Ellipse))        return ShapeKind.Ellipse;
        if (v.Equals(A.ShapeTypeValues.Triangle))       return ShapeKind.Triangle;
        if (v.Equals(A.ShapeTypeValues.Diamond))        return ShapeKind.Diamond;
        return ShapeKind.Other;
    }

    private static double CalcCornerRadius(ShapeKind kind, A.PresetGeometry? geom, double w, double h)
    {
        if (kind != ShapeKind.RoundRect) return 0;
        var adj = geom?.AdjustValueList?
            .Elements<A.ShapeGuide>()
            .FirstOrDefault(g => g.Name?.Value == "adj");
        double pct = 16667;
        if (adj?.Formula?.Value is string f && f.StartsWith("val "))
            double.TryParse(f[4..], out pct);
        return Math.Min(w, h) * pct / 200000.0;
    }

    // ─── Fill / Stroke ────────────────────────────────────────────────────────

    private WColor? ParseFill(P.ShapeProperties spPr)
    {
        if (spPr.GetFirstChild<A.NoFill>() != null) return null;
        var sf = spPr.GetFirstChild<A.SolidFill>();
        return sf != null ? ParseColor(sf) : null;
    }

    private (WColor? color, double widthPt) ParseStroke(A.Outline? ol)
    {
        if (ol == null) return (null, 0);
        if (ol.GetFirstChild<A.NoFill>() != null) return (null, 0);
        double w  = (ol.Width?.Value ?? 12700) * EmuToPt;
        var sf    = ol.GetFirstChild<A.SolidFill>();
        return (sf != null ? ParseColor(sf) : WColor.FromRgb(0, 0, 0), w);
    }

    private WColor? ParseColor(A.SolidFill fill)
    {
        var rgb = fill.GetFirstChild<A.RgbColorModelHex>();
        if (rgb?.Val?.Value is string hex) return ParseHex(hex);

        var sys = fill.GetFirstChild<A.SystemColor>();
        if (sys?.LastColor?.Value is string lc) return ParseHex(lc);

        var scheme = fill.GetFirstChild<A.SchemeColor>();
        if (scheme?.Val?.Value.ToString() is string key && _themeColors.TryGetValue(key, out var tc))
        {
            var alpha = scheme.GetFirstChild<A.Alpha>()?.Val?.Value;
            if (alpha.HasValue)
                return WColor.FromArgb((byte)(255 * alpha.Value / 100000), tc.R, tc.G, tc.B);
            return tc;
        }

        return null;
    }

    // ─── Text ─────────────────────────────────────────────────────────────────

    private TextInfo ParseTextBody(P.TextBody txBody)
    {
        var anchor = txBody.BodyProperties?.Anchor?.Value;
        VTextAlign vAlign;
        if (anchor?.Equals(A.TextAnchoringTypeValues.Top) == true)         vAlign = VTextAlign.Top;
        else if (anchor?.Equals(A.TextAnchoringTypeValues.Bottom) == true) vAlign = VTextAlign.Bottom;
        else                                                                vAlign = VTextAlign.Center;

        var paras = new List<TextParagraph>();
        foreach (var para in txBody.Elements<A.Paragraph>())
        {
            var algn = para.ParagraphProperties?.Alignment?.Value;
            HTextAlign hAlign;
            if (algn?.Equals(A.TextAlignmentTypeValues.Left) == true)      hAlign = HTextAlign.Left;
            else if (algn?.Equals(A.TextAlignmentTypeValues.Right) == true) hAlign = HTextAlign.Right;
            else if (algn?.Equals(A.TextAlignmentTypeValues.Justified) == true) hAlign = HTextAlign.Justify;
            else hAlign = HTextAlign.Center;

            var runs = new List<TextRunInfo>();
            foreach (var run in para.Elements<A.Run>())
            {
                var t = run.Text?.Text;
                if (string.IsNullOrEmpty(t)) continue;
                var rPr    = run.RunProperties;
                double sz  = (rPr?.FontSize?.Value ?? 1400) / 100.0;
                bool bold  = rPr?.Bold?.Value == true;
                bool ital  = rPr?.Italic?.Value == true;
                WColor? fc = null;
                var sf     = rPr?.GetFirstChild<A.SolidFill>();
                if (sf != null) fc = ParseColor(sf);
                runs.Add(new TextRunInfo(t, sz, bold, ital, fc));
            }
            if (runs.Count > 0)
                paras.Add(new TextParagraph(runs, hAlign));
        }

        return new TextInfo(paras, vAlign);
    }

    // ─── Helpers ──────────────────────────────────────────────────────────────

    private static WColor ParseHex(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length >= 6)
            return WColor.FromRgb(
                Convert.ToByte(hex[..2], 16),
                Convert.ToByte(hex[2..4], 16),
                Convert.ToByte(hex[4..6], 16));
        return WColor.FromRgb(0, 0, 0);
    }
}
