using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using wpf_pptx_editor.Forms;
using wpf_pptx_editor.Forms.Services;
using wpf_pptx_editor.Forms.UI.Views;
using wpf_pptx_editor.Forms.ViewModels;

namespace wpf_pptx_editor;

public class App : Application
{
    private IHost? _host;

    public static IServiceProvider? ServiceProvider { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _host = Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                ConfigureServices(services);
            })
            .Build();

        ServiceProvider = _host.Services;
        AppServices.Current = _host.Services;

        var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
        mainWindow.Icon = new System.Windows.Media.Imaging.BitmapImage(
            new Uri("pack://application:,,,/wpf_pptx_editor;component/app.ico"));
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _host?.Dispose();
        base.OnExit(e);
    }

    private void ConfigureServices(IServiceCollection services)
    {
        services.AddSingleton<IPptxService, PptxService>();
        services.AddSingleton<MainWindowViewModel>();
        services.AddSingleton<MainWindow>();
    }
    // PptxWriter, SlideEditorViewModel은 MainWindowViewModel 내부에서 직접 생성
}
