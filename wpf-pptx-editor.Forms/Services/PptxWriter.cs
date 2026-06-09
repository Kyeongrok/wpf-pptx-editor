using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using wpf_pptx_editor.Forms.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace wpf_pptx_editor.Forms.Services;

public class PptxWriter
{
    private const long DipToEmu = 9525L; // 914400 / 96

    // ── 새 파일 생성 ──────────────────────────────────────────────────────────

    public void CreateNew(string filePath)
    {
        using var doc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);

        var presPart = doc.AddPresentationPart();
        presPart.Presentation = new P.Presentation();

        // 테마
        var themePart = presPart.AddNewPart<ThemePart>();
        themePart.Theme = BuildMinimalTheme();

        // 슬라이드 마스터
        var masterPart = presPart.AddNewPart<SlideMasterPart>();
        masterPart.SlideMaster = BuildMinimalSlideMaster();
        masterPart.AddPart(themePart);

        // 슬라이드 레이아웃
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = BuildMinimalSlideLayout();
        layoutPart.AddPart(masterPart);
        masterPart.SlideMaster.SlideLayoutIdList = new P.SlideLayoutIdList(
            new P.SlideLayoutId { Id = 2147483649U, RelationshipId = masterPart.GetIdOfPart(layoutPart) });

        // 프레젠테이션 설정
        presPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(
            new P.SlideMasterId { Id = 2147483648U, RelationshipId = presPart.GetIdOfPart(masterPart) });
        presPart.Presentation.SlideIdList = new P.SlideIdList();
        presPart.Presentation.SlideSize = new P.SlideSize { Cx = 9144000, Cy = 5143500 };
        presPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };

        // 첫 슬라이드
        AddBlankSlide(doc, layoutPart);

        presPart.Presentation.Save();
    }

    // ── 슬라이드 저장 ─────────────────────────────────────────────────────────

    public void SaveSlide(string filePath, int slideIndex, IList<EditableShape> shapes)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presPart = doc.PresentationPart!;

        var slideId = presPart.Presentation.SlideIdList!
            .Elements<P.SlideId>().ElementAtOrDefault(slideIndex);
        if (slideId == null) return;

        var slidePart = (SlidePart)presPart.GetPartById(slideId.RelationshipId!.Value!);
        var spTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        // 기존 도형 제거 (그룹 속성 노드는 유지)
        spTree.Elements<P.Shape>().ToList().ForEach(s => s.Remove());
        spTree.Elements<P.ConnectionShape>().ToList().ForEach(c => c.Remove());

        uint id = 2;
        foreach (var shape in shapes)
            spTree.Append(BuildShape(shape, id++));

        slidePart.Slide.Save();
        presPart.Presentation.Save();
    }

    // ── 슬라이드 추가 ─────────────────────────────────────────────────────────

    public void AddSlide(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var layoutPart = doc.PresentationPart!
            .SlideMasterParts.First()
            .SlideLayoutParts.First();
        AddBlankSlide(doc, layoutPart);
        doc.PresentationPart!.Presentation.Save();
    }

    // ── 내부 헬퍼 ────────────────────────────────────────────────────────────

    private static void AddBlankSlide(PresentationDocument doc, SlideLayoutPart layoutPart)
    {
        var presPart = doc.PresentationPart!;
        var slidePart = presPart.AddNewPart<SlidePart>();
        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 },
                    new A.ChildOffset { X = 0, Y = 0 },
                    new A.ChildExtents { Cx = 0, Cy = 0 })))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        slidePart.AddPart(layoutPart);

        uint maxId = presPart.Presentation.SlideIdList!
            .Elements<P.SlideId>().Select(s => s.Id!.Value).DefaultIfEmpty(255U).Max();
        presPart.Presentation.SlideIdList.Append(
            new P.SlideId { Id = maxId + 1, RelationshipId = presPart.GetIdOfPart(slidePart) });

        slidePart.Slide.Save();
    }

    private static OpenXmlElement BuildShape(EditableShape s, uint id)
    {
        long x  = (long)(s.X * DipToEmu);
        long y  = (long)(s.Y * DipToEmu);
        long cx = (long)(s.Width * DipToEmu);
        long cy = (long)(s.Height * DipToEmu);
        int  sw = (int)(s.StrokeWidthPt * 12700);
        bool isLine = s.Kind == ShapeKind.Line;

        if (isLine)
            return BuildConnectionShape(s, id, x, y, cx, cy, sw);

        string prst = s.Kind switch
        {
            ShapeKind.RoundRect => "roundRect",
            ShapeKind.Ellipse   => "ellipse",
            ShapeKind.Triangle  => "triangle",
            ShapeKind.Diamond   => "diamond",
            _                   => "rect"
        };

        var xfrm = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = cx, Cy = cy });
        if (s.FlipH) xfrm.HorizontalFlip = true;
        if (s.FlipV) xfrm.VerticalFlip = true;

        var outlineEl = new A.Outline(
            new A.SolidFill(new A.RgbColorModelHex { Val = ToHex(s.StrokeColor) }))
            { Width = Math.Max(sw, 1) };

        OpenXmlElement fillEl = s.FillColor.A == 0
            ? (OpenXmlElement)new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = ToHex(s.FillColor) });
        var spPr = new P.ShapeProperties(
            xfrm,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(prst) },
            fillEl,
            outlineEl);

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Shape {id}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            spPr,
            BuildTextBody(s));
    }

    private static P.ConnectionShape BuildConnectionShape(EditableShape s, uint id,
        long x, long y, long cx, long cy, int sw)
    {
        var xfrm = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = cx, Cy = cy });
        if (s.FlipH) xfrm.HorizontalFlip = true;
        if (s.FlipV) xfrm.VerticalFlip = true;

        var ln = new A.Outline(
            new A.SolidFill(new A.RgbColorModelHex { Val = ToHex(s.StrokeColor) }),
            new A.PresetDash { Val = A.PresetLineDashValues.Solid },
            new A.Round(),
            new A.HeadEnd { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium },
            new A.TailEnd { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium })
        {
            Width = Math.Max(sw, 1),
            CapType = A.LineCapValues.Flat,
            CompoundLineType = A.CompoundLineValues.Single
        };

        var spPr = new P.ShapeProperties(
            xfrm,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.StraightConnector1 },
            new A.NoFill(),
            ln);

        return new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Connector {id}" },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            spPr);
    }

    // (kept for forward-compat reference — no longer used)
    private static P.ShapeStyle BuildLineStyle() =>
        new P.ShapeStyle(
            new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 1U },
            new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
            new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
            new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Dark1 })
                { Index = A.FontCollectionIndexValues.Minor });

    private static P.TextBody BuildTextBody(EditableShape s)
    {
        var bodyPr = new A.BodyProperties();
        bodyPr.Anchor = s.VAlign switch
        {
            VTextAlign.Top    => new A.TextAnchoringTypeValues("t"),
            VTextAlign.Bottom => new A.TextAnchoringTypeValues("b"),
            _                 => new A.TextAnchoringTypeValues("ctr")
        };

        A.Paragraph para;
        if (!string.IsNullOrEmpty(s.Text))
        {
            var pPr = new A.ParagraphProperties();
            pPr.Alignment = s.HAlign switch
            {
                HTextAlign.Left    => new A.TextAlignmentTypeValues("l"),
                HTextAlign.Right   => new A.TextAlignmentTypeValues("r"),
                HTextAlign.Justify => new A.TextAlignmentTypeValues("just"),
                _                  => new A.TextAlignmentTypeValues("ctr")
            };
            var rPr = new A.RunProperties { Language = "ko-KR", FontSize = (int)(s.FontSizePt * 100), Dirty = false };
            para = new A.Paragraph(pPr, new A.Run(rPr, new A.Text(s.Text)));
        }
        else
        {
            para = new A.Paragraph();
        }

        return new P.TextBody(bodyPr, new A.ListStyle(), para);
    }

    private static string ToHex(System.Windows.Media.Color c)
        => $"{c.R:X2}{c.G:X2}{c.B:X2}";

    // ── 최소 PPTX 구조 빌더 ───────────────────────────────────────────────────

    private static A.Theme BuildMinimalTheme()
    {
        return new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = new A.SystemColorValues("windowText"), LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = new A.SystemColorValues("window"), LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "A9D18E" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" }))
                { Name = "Office" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont { Typeface = "+mj-lt" },
                        new A.EastAsianFont { Typeface = "+mj-ea" },
                        new A.ComplexScriptFont { Typeface = "+mj-cs" },
                        new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" }),
                    new A.MinorFont(
                        new A.LatinFont { Typeface = "+mn-lt" },
                        new A.EastAsianFont { Typeface = "+mn-ea" },
                        new A.ComplexScriptFont { Typeface = "+mn-cs" },
                        new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" }))
                { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") }),
                        new A.GradientFill(new A.GradientStopList(), new A.LinearGradientFill { Angle = 5400000, Scaled = false }),
                        new A.GradientFill(new A.GradientStopList(), new A.LinearGradientFill { Angle = 5400000, Scaled = false })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") })) { Width = 6350 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") })) { Width = 12700 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") })) { Width = 19050 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") }),
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues("phClr") }),
                        new A.GradientFill(new A.GradientStopList(), new A.LinearGradientFill { Angle = 5400000, Scaled = false })))
                { Name = "Office" })) { Name = "Office Theme" };
    }

    private static P.SlideMaster BuildMinimalSlideMaster()
    {
        var defRPr = () => new A.DefaultRunProperties { Language = "ko-KR" };
        var lvl1 = () => new A.Level1ParagraphProperties(defRPr());

        return new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 },
                    new A.ChildOffset { X = 0, Y = 0 },
                    new A.ChildExtents { Cx = 0, Cy = 0 })))),
            new P.ColorMap
            {
                Background1 = new A.ColorSchemeIndexValues("lt1"),
                Text1 = new A.ColorSchemeIndexValues("dk1"),
                Background2 = new A.ColorSchemeIndexValues("lt2"),
                Text2 = new A.ColorSchemeIndexValues("dk2"),
                Accent1 = new A.ColorSchemeIndexValues("accent1"),
                Accent2 = new A.ColorSchemeIndexValues("accent2"),
                Accent3 = new A.ColorSchemeIndexValues("accent3"),
                Accent4 = new A.ColorSchemeIndexValues("accent4"),
                Accent5 = new A.ColorSchemeIndexValues("accent5"),
                Accent6 = new A.ColorSchemeIndexValues("accent6"),
                Hyperlink = new A.ColorSchemeIndexValues("hlink"),
                FollowedHyperlink = new A.ColorSchemeIndexValues("folHlink")
            },
            new P.TextStyles(
                new P.TitleStyle(lvl1()),
                new P.BodyStyle(lvl1()),
                new P.OtherStyle(lvl1())));
    }

    private static P.SlideLayout BuildMinimalSlideLayout()
    {
        return new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 },
                    new A.ChildOffset { X = 0, Y = 0 },
                    new A.ChildExtents { Cx = 0, Cy = 0 })))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = new P.SlideLayoutValues("blank"), Preserve = true };
    }
}
