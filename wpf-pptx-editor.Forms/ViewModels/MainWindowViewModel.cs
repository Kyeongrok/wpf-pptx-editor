using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using wpf_pptx_editor.Forms.Models;
using wpf_pptx_editor.Forms.Services;

namespace wpf_pptx_editor.Forms.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private readonly IPptxService _pptxService;
    private readonly PptxWriter _writer = new();

    private string _currentFilePath = "";
    private int _currentSlideIndex = -1;

    [ObservableProperty] private ObservableCollection<SlideItemViewModel> _slides = new();
    [ObservableProperty] private SlideItemViewModel? _selectedSlide;
    [ObservableProperty] private bool _isLoading;
    [ObservableProperty] private string _windowTitle = "PPTX Editor";
    [ObservableProperty] private SlideEditorViewModel _editor = new();
    [ObservableProperty] private string _statusMessage = "준비";

    public MainWindowViewModel(IPptxService pptxService)
    {
        _pptxService = pptxService;
        InitializeEmptyPresentation();
    }

    private void InitializeEmptyPresentation()
    {
        Slides.Clear();
        Editor.LoadFromSlideInfo(new SlideInfo(960, 540, Array.Empty<ShapeInfo>()));
        _currentSlideIndex = 0;
        Slides.Add(new SlideItemViewModel(1, null));
        SelectedSlide = Slides[0];
    }

    // ── 파일 열기 ────────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task OpenFileAsync()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "PowerPoint Files|*.pptx|All Files|*.*",
            Title = "PPTX 파일 열기"
        };
        if (dlg.ShowDialog() != true) return;

        _currentFilePath = dlg.FileName;
        IsLoading = true;
        Slides.Clear();
        SelectedSlide = null;

        try
        {
            var infos = await Task.Run(() => _pptxService.LoadPresentation(_currentFilePath));
            foreach (var s in infos) Slides.Add(s);
            WindowTitle = $"PPTX Editor - {Path.GetFileName(_currentFilePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"파일 열기 실패:\n{ex.Message}", "오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { IsLoading = false; }

        if (Slides.Count > 0) SelectedSlide = Slides[0];
    }

    // ── 새 파일 ──────────────────────────────────────────────────────────────

    [RelayCommand]
    private void NewFile()
    {
        _currentFilePath = "";
        WindowTitle = "PPTX Editor";
        InitializeEmptyPresentation();
    }

    // ── 저장 ─────────────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task SaveAsync()
    {
        if (string.IsNullOrEmpty(_currentFilePath))
        {
            await SaveAsAsync();
            return;
        }
        await DoSaveAsync(_currentFilePath);
    }

    [RelayCommand]
    private async Task SaveAsAsync()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "PowerPoint Files|*.pptx",
            Title = "다른 이름으로 저장",
            FileName = Path.GetFileName(_currentFilePath)
        };
        if (dlg.ShowDialog() != true) return;

        // 새 경로에 파일이 없으면 새로 생성
        if (!File.Exists(dlg.FileName))
            _writer.CreateNew(dlg.FileName);

        _currentFilePath = dlg.FileName;
        WindowTitle = $"PPTX Editor - {Path.GetFileName(_currentFilePath)}";
        await DoSaveAsync(_currentFilePath);
    }

    private async Task DoSaveAsync(string path)
    {
        if (_currentSlideIndex < 0) return;
        IsLoading = true;
        try
        {
            var shapes = Editor.Shapes.ToList();
            int idx = _currentSlideIndex;
            await Task.Run(() => _writer.SaveSlide(path, idx, shapes));

            // 썸네일 갱신
            RefreshCurrentThumbnail();
            StatusMessage = $"저장 완료  {Path.GetFileName(path)}";
            _ = ClearStatusAfterDelayAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 실패:\n{ex.Message}", "오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { IsLoading = false; }
    }

    // ── 슬라이드 추가 ────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task AddSlideAsync()
    {
        try
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                await Task.Run(() => _writer.AddSlide(_currentFilePath));

            int newIndex = Slides.Count;
            Slides.Add(new SlideItemViewModel(newIndex + 1, null));
            SelectedSlide = Slides[newIndex];
        }
        catch (Exception ex)
        {
            MessageBox.Show($"슬라이드 추가 실패:\n{ex.Message}", "오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ── 슬라이드 선택 ────────────────────────────────────────────────────────

    partial void OnSelectedSlideChanged(SlideItemViewModel? value)
    {
        if (value == null) return;

        // 이전 슬라이드 자동 저장 (편집 내용 보존)
        if (_currentSlideIndex >= 0 && !string.IsNullOrEmpty(_currentFilePath))
        {
            var shapes = Editor.Shapes.ToList();
            int prevIdx = _currentSlideIndex;
            Task.Run(() =>
            {
                try { _writer.SaveSlide(_currentFilePath, prevIdx, shapes); }
                catch { /* 자동 저장 실패는 무시 */ }
            });
        }

        _currentSlideIndex = value.SlideIndex;
        _ = LoadSlideForEditAsync(value.SlideIndex);
    }

    private async Task LoadSlideForEditAsync(int slideIndex)
    {
        if (string.IsNullOrEmpty(_currentFilePath))
        {
            Editor.LoadFromSlideInfo(new SlideInfo(960, 540, Array.Empty<ShapeInfo>()));
            return;
        }
        IsLoading = true;
        try
        {
            var slideInfo = await Task.Run(() => _pptxService.GetSlideInfo(slideIndex));
            Editor.LoadFromSlideInfo(slideInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"슬라이드 로드 실패:\n{ex.Message}", "오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { IsLoading = false; }
    }

    private async Task ClearStatusAfterDelayAsync()
    {
        await Task.Delay(3000);
        StatusMessage = "준비";
    }

    // ── 썸네일 갱신 ──────────────────────────────────────────────────────────

    private void RefreshCurrentThumbnail()
    {
        if (_currentSlideIndex < 0 || _currentSlideIndex >= Slides.Count) return;
        var slideInfo = Editor.ToSlideInfo();
        Application.Current.Dispatcher.Invoke(() =>
        {
            var renderer = new SlideRenderer();
            var bmp = renderer.RenderToBitmap(slideInfo, 240);
            Slides[_currentSlideIndex].Thumbnail = bmp;
        });
    }
}
