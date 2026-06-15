using System.Runtime.InteropServices;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using PDFly.Models;
using PDFly.Services;
using PDFly.ViewModels;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;

namespace PDFly;

public sealed partial class MainWindow : Window
{
    public MainViewModel ViewModel { get; }

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon("pdfly.ico");

        // Match the old WPF build's default footprint (540×660 DIPs), centred.
        SizeAndCenter(540, 660);

        ViewModel = new MainViewModel(PdflySettings.Load());
        Root.DataContext = ViewModel;

        AppWindow.Closing += OnAppWindowClosing;
        Closed += OnWindowClosed;
    }

    /// <summary>Size + centre the window in device-independent pixels. AppWindow.Resize
    /// works in physical pixels, so we scale by the monitor DPI first.</summary>
    private void SizeAndCenter(int widthDip, int heightDip)
    {
        try
        {
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            uint dpi = GetDpiForWindow(hwnd);
            if (dpi == 0) dpi = 96;
            double scale = dpi / 96.0;
            int w = (int)(widthDip * scale);
            int h = (int)(heightDip * scale);

            var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary);
            var work = area.WorkArea;
            int x = work.X + Math.Max(0, (work.Width - w) / 2);
            int y = work.Y + Math.Max(0, (work.Height - h) / 2);
            AppWindow.MoveAndResize(new RectInt32(x, y, w, h));
        }
        catch { /* fall back to whatever default size WinUI picks */ }
    }

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    /// <summary>Queue files passed in via command line ("Open with PDFly" / drop onto exe).</summary>
    public void QueuePaths(IEnumerable<string> paths) => ViewModel.AddPaths(paths);

    // ----------------------------------------------------------- drag & drop

    private void OnDragOver(object sender, DragEventArgs e)
    {
        if (e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
            ViewModel.IsDragOver = true;
            UpdateDropZone(active: true);
            if (e.DragUIOverride is { } overlay)
            {
                overlay.Caption = "Add to PDFly";
                overlay.IsCaptionVisible = true;
                overlay.IsGlyphVisible = true;
            }
        }
        else
        {
            e.AcceptedOperation = DataPackageOperation.None;
        }
    }

    private void OnDragLeave(object sender, DragEventArgs e)
    {
        ViewModel.IsDragOver = false;
        UpdateDropZone(active: false);
    }

    private async void OnDrop(object sender, DragEventArgs e)
    {
        ViewModel.IsDragOver = false;
        UpdateDropZone(active: false);
        if (!e.DataView.Contains(StandardDataFormats.StorageItems)) return;

        var deferral = e.GetDeferral();
        try
        {
            var items = await e.DataView.GetStorageItemsAsync();
            ViewModel.AddPaths(items.Select(i => i.Path));
        }
        catch { /* malformed drop payload — ignore */ }
        finally { deferral.Complete(); }
    }

    private void UpdateDropZone(bool active)
    {
        var key = active ? "AccentFillColorDefaultBrush" : "TextFillColorSecondaryBrush";
        if (Application.Current.Resources[key] is Brush b)
            DropZoneOutline.Stroke = b;
        DropZoneOutline.StrokeThickness = active ? 2.0 : 1.3;
    }

    // ----------------------------------------------------- list interactions

    private void OnRowDoubleClick(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (ItemOf(sender) is { } item) ViewModel.OpenItem(item);
    }

    private void OnOpenItem(object sender, RoutedEventArgs e)    { if (ItemOf(sender) is { } i) ViewModel.OpenItem(i); }
    private void OnRevealItem(object sender, RoutedEventArgs e)  { if (ItemOf(sender) is { } i) ViewModel.RevealItem(i); }
    private void OnRequeueItem(object sender, RoutedEventArgs e) { if (ItemOf(sender) is { } i) ViewModel.RequeueItem(i); }
    private void OnRemoveItem(object sender, RoutedEventArgs e)  { if (ItemOf(sender) is { } i) ViewModel.RemoveItem(i); }

    /// <summary>The double-tap handler receives a Grid whose DataContext is the item;
    /// MenuFlyoutItem clicks come back with the item carried in <c>Tag</c> (set via
    /// <c>Tag="{x:Bind}"</c> in the template). Try both, take whichever resolves.</summary>
    private static ConversionItem? ItemOf(object sender)
    {
        if (sender is not FrameworkElement fe) return null;
        return fe.DataContext as ConversionItem ?? fe.Tag as ConversionItem;
    }

    // ------------------------------------------------------------- lifecycle

    private bool _forceClose;
    private async void OnAppWindowClosing(
        Microsoft.UI.Windowing.AppWindow sender,
        Microsoft.UI.Windowing.AppWindowClosingEventArgs args)
    {
        if (_forceClose) return;
        if (!ViewModel.HasActiveItems) return;

        args.Cancel = true;
        int n = ViewModel.Items.Count(i => i.IsActive);
        var dialog = new ContentDialog
        {
            Title = "PDFly",
            Content = $"PDFly is still converting {n} file{(n == 1 ? string.Empty : "s")}.\n\nClose anyway?",
            PrimaryButtonText = "Close",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = Content.XamlRoot,
        };
        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            _forceClose = true;
            Close();
        }
    }

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        ViewModel.Dispose();
    }
}
