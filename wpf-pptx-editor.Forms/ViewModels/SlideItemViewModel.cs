using System.Windows.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;

namespace wpf_pptx_editor.Forms.ViewModels;

public partial class SlideItemViewModel : ObservableObject
{
    public int SlideNumber { get; }
    public int SlideIndex { get; }

    [ObservableProperty]
    private BitmapImage? _thumbnail;

    public SlideItemViewModel(int slideNumber, BitmapImage? thumbnail = null)
    {
        SlideNumber = slideNumber;
        SlideIndex = slideNumber - 1;
        _thumbnail = thumbnail;
    }
}
