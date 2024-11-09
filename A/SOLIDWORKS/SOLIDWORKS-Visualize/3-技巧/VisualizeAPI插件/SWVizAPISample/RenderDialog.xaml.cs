using System.ComponentModel;
using System.Windows;

namespace SWVizAPISample
{
    public partial class RenderDialog : Window
    {
        public RenderDialogViewModel ViewModel { get; }

        public RenderDialog(string message)
        {
            InitializeComponent();
            ViewModel = new RenderDialogViewModel(this);
            DataContext = ViewModel;
        }

        private void RenderDialog_OnClosing(object sender, CancelEventArgs e)
        {
            ViewModel.DialogCancel = true;
        }
    }
}
