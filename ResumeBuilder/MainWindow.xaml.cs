using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ResumeBuilder.ViewModels;

namespace ResumeBuilder;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void ApiKeyBox_OnPasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm && sender is PasswordBox box)
        {
            vm.ApiKeyInput = box.Password;
        }
    }

    private void TitleBar_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
        {
            return;
        }

        if (e.OriginalSource is DependencyObject source && IsWithinInteractiveControl(source))
        {
            return;
        }

        if (e.ClickCount == 2 && ResizeMode != ResizeMode.NoResize)
        {
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
            return;
        }

        try
        {
            DragMove();
        }
        catch
        {
            // Ignore drag exceptions when interaction state changes mid-drag.
        }
    }

    private static bool IsWithinInteractiveControl(DependencyObject source)
    {
        DependencyObject? current = source;
        while (current != null)
        {
            if (current is Button || current is TextBox || current is PasswordBox || current is ScrollBar)
            {
                return true;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private void TitleBarMinimize_OnClick(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void TitleBarMaximize_OnClick(object sender, RoutedEventArgs e)
    {
        if (ResizeMode == ResizeMode.NoResize)
        {
            return;
        }

        WindowState = WindowState.Maximized;
    }

    private void TitleBarRestore_OnClick(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Normal;
    }

    private void TitleBarClose_OnClick(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
