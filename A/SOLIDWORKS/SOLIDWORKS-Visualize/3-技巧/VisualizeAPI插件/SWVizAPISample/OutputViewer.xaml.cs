using System.Globalization;
using System;
using System.Windows;
using System.Windows.Data;

namespace SWVizAPISample
{
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isVisible && isVisible)
                return Visibility.Visible;
            else
                return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }   
    public partial class OutputViewerDialog : Window
    {
        public OutputViewerDialogViewModel ViewModel { get; }

        public OutputViewerDialog(string message)
        {
            InitializeComponent();

            ViewModel = new OutputViewerDialogViewModel(this);
            DataContext = ViewModel;
        }
    }
}
