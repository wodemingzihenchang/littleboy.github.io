using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.Runtime.CompilerServices;
using SolidWorks.Visualize.Interfaces;

namespace SWVizAPISample
{
    public class RenderDialogViewModel : INotifyPropertyChanged
    {
        private RenderDialog _parentWindow;
        private string _outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private ImageFormat_e _selectedImageFormat;
        private string _jobName = "Default";
        private bool _includeAlpha;
        private float _sizeX = 1920.0f;
        private float _sizeY = 1080.0f;
        private bool _accurateRender;
        private bool _terminationModeTime;
        private bool _terminationModeQuality;
        private int _renderPasses = 1000;
        private TimeSpan _renderDuration = new TimeSpan(0,0,30);
        private bool _enableDenoiser = false;
        private bool _dialogCancel = false;
        public string OutputFolder
        {
            get { return _outputFolder; }
            set
            {
                if (_outputFolder != value)
                {
                    _outputFolder = value;
                    OnPropertyChanged();
                }
            }
        }
        public ImageFormat_e SelectedImageFormat
        {
            get { return _selectedImageFormat; }
            set
            {
                if (_selectedImageFormat != value)
                {
                    _selectedImageFormat = value;
                    OnPropertyChanged();
                }
            }
        }
        public IEnumerable<ImageFormat_e> ImageFormatOptions
        {
            get { return (IEnumerable<ImageFormat_e>)Enum.GetValues(typeof(ImageFormat_e)); }
        }

        public string JobName
        {
            get { return _jobName; }
            set
            {
                if (_jobName != value)
                {
                    _jobName = value;
                    OnPropertyChanged();
                }
            }
        }
        public int RenderPasses
        {
            get { return _renderPasses; }
            set
            {
                if (_renderPasses != value)
                {
                    _renderPasses = value;
                    OnPropertyChanged();
                }
            }
        }

        public TimeSpan TimeLimit
        {
            get { return _renderDuration; }
            set
            {
                if (_renderDuration != value)
                {
                    _renderDuration = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IncludeAlpha
        {
            get { return _includeAlpha; }
            set
            {
                if (_includeAlpha != value)
                {
                    _includeAlpha = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool AccurateRender
        {
            get { return _accurateRender; }
            set
            {
                if (_accurateRender != value)
                {
                    _accurateRender = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool QualityMode
        {
            get { return _terminationModeQuality; }
            set
            {
                if (_terminationModeQuality != value)
                {
                    _terminationModeQuality = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool TimeMode
        {
            get { return _terminationModeTime; }
            set
            {
                if (_terminationModeTime != value)
                {
                    _terminationModeTime = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool EnableDenoiser
        {
            get { return _enableDenoiser; }
            set
            {
                if (_enableDenoiser != value)
                {
                    _enableDenoiser = value;
                    OnPropertyChanged();
                }
            }
        }
        public float Width
        {
            get { return _sizeX; }
            set
            {
                if (_sizeX != value)
                {
                    _sizeX = value;
                    OnPropertyChanged();
                }
            }
        }
        public float Height
        {
            get { return _sizeY; }
            set
            {
                if (_sizeY != value)
                {
                    _sizeY = value;
                    OnPropertyChanged();
                }
            }
        }
        // Add properties for other UI elements (sizeX, sizeY, etc.) in a similar manner

        public RelayCommand OkCommand { get; }
        public RelayCommand CancelCommand { get; }
        public RelayCommand BrowseCommand { get; }

        public bool DialogCancel
        {
            get => _dialogCancel;
            set => _dialogCancel = value;
        }

        public RenderDialogViewModel(RenderDialog parentWindow)
        {
            _parentWindow = parentWindow;
            OkCommand = new RelayCommand(OkCommandExecute);
            CancelCommand = new RelayCommand(CancelCommandExecute);
            BrowseCommand = new RelayCommand(BrowseCommandExecute);
            SelectedImageFormat = ImageFormat_e.JPEG;
            AccurateRender = true;
            QualityMode = true;
        }

        private void OkCommandExecute(object obj)
        {
            _parentWindow?.Close();
            DialogCancel = false;
        }

        private void CancelCommandExecute(object obj)
        {
            // Perform any cleanup or additional logic here
            DialogCancel = true;
            // Close the dialog window
            _parentWindow?.Close();
        }

        private void BrowseCommandExecute(object obj)
        {
            // Create a FolderBrowserDialog
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();

            // Show the dialog and get the result
            var result = folderDialog.ShowDialog();

            // If user selects a folder, update the OutputFolder property with the selected folder path
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                OutputFolder = folderDialog.SelectedPath;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
