using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace Demo;

public partial class DatePickerWindow : Window
{
    public DatePickerWindow()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}