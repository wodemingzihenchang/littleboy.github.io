using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace SWVizAPISample
{
    public class OutputViewerDialogViewModel : INotifyPropertyChanged
    {
        public event Action RenderCanceled;
        public event Action RenderCanceledAndSave;

        private OutputViewerDialog _parentWindow;
        private TimeSpan _timeSpan = TimeSpan.Zero;
        private int _framesRendered = 0;

        private double _renderProgress = 0;
        public double RenderProgress
        {
            get { return _renderProgress; }
            set { _renderProgress = value; OnPropertyChanged(); }
        }

        private BitmapImage _renderedImage;
        public BitmapImage RenderedImage
        {
            get { return _renderedImage; }
            set
            {
                _renderedImage = value;
                
                OnPropertyChanged();
            }
        }

        public ICommand CancelCommand { get; }
        public ICommand CancelAndSaveCommand { get; }
        public ICommand CloseDialogCommand { get; }
        
        public TimeSpan TimeSpan
        {
            get => _timeSpan;
            set { _timeSpan = value; OnPropertyChanged(); }
        }

        public int FramesRendered
        {
            get => _framesRendered;
            set
            {
                _framesRendered = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FramesInfo));
            }
        }

        public string FramesInfo
            => TotalFrames >= 0 ? $"{FramesRendered} / {TotalFrames}" : $"{FramesRendered}";

        public int TotalFrames { get; set; }

        public float AspectRatio { get; set; }

        public OutputViewerDialogViewModel(OutputViewerDialog parentWindow)
        {
            // Initialize commands
            _parentWindow = parentWindow;
            CancelCommand = new RelayCommand(Cancel);
            CancelAndSaveCommand = new RelayCommand(CancelAndSave);
            CloseDialogCommand = new RelayCommand(CloseDialog);
        }

        private void Cancel(object obj)
        {
            RenderCanceled?.Invoke();
            CloseDialog(null);
        }

        private void CancelAndSave(object obj)
        {
            RenderCanceledAndSave?.Invoke();
            CloseDialog(null);
        }
        private void CloseDialog(object obj)
        {
            _parentWindow?.Close();
            Cleanup();
        }

        private void Cleanup()
        {
            if (RenderedImage != null)
            {
                RenderedImage.StreamSource.Dispose();
                RenderedImage = null;
            }
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
