using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml.Packaging;
using wpf_pptx_editor.Forms.Models;
using wpf_pptx_editor.Forms.ViewModels;

namespace wpf_pptx_editor.Forms.Services;

public interface IPptxService : IDisposable
{
    IReadOnlyList<SlideItemViewModel> LoadPresentation(string filePath);
    BitmapImage GetSlideImage(int slideIndex);
    SlideInfo GetSlideInfo(int slideIndex);
}

public sealed class PptxService : IPptxService
{
    private readonly PptxParser _parser = new();
    private readonly SlideRenderer _renderer = new();
    private PresentationDocument? _doc;
    private readonly List<SlideInfo> _slideInfos = new();

    public IReadOnlyList<SlideItemViewModel> LoadPresentation(string filePath)
    {
        _doc?.Dispose();
        _slideInfos.Clear();

        _doc = PresentationDocument.Open(filePath, false);
        _parser.LoadThemeColors(_doc);

        int count = _parser.GetSlideCount(_doc);
        for (int i = 0; i < count; i++)
            _slideInfos.Add(_parser.ParseSlide(_doc, i));

        // Render thumbnails on UI thread via Dispatcher
        var result = new List<SlideItemViewModel>();
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            for (int i = 0; i < _slideInfos.Count; i++)
            {
                var thumbnail = _renderer.RenderToBitmap(_slideInfos[i], 240);
                result.Add(new SlideItemViewModel(i + 1, thumbnail));
            }
        });
        return result;
    }

    public SlideInfo GetSlideInfo(int slideIndex)
    {
        if (slideIndex < 0 || slideIndex >= _slideInfos.Count)
            throw new ArgumentOutOfRangeException(nameof(slideIndex));
        return _slideInfos[slideIndex];
    }

    public BitmapImage GetSlideImage(int slideIndex)
    {
        if (slideIndex < 0 || slideIndex >= _slideInfos.Count)
            throw new ArgumentOutOfRangeException(nameof(slideIndex));

        BitmapImage? image = null;
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
            image = _renderer.RenderToBitmap(_slideInfos[slideIndex], 1920));
        return image!;
    }

    public void Dispose()
    {
        _doc?.Dispose();
        _slideInfos.Clear();
    }
}
