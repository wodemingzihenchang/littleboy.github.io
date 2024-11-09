using System.IO;
using System.Timers;
using System.Windows.Media.Imaging;
using SolidWorks.Visualize.Interfaces;

namespace SWVizAPISample
{
    public class RenderHandler : IRenderNotificationHandler
    {
        private readonly Timer _requestPreviewTimer;

        private bool _previewRequestPending = false;
        
        public event ChangeRenderActionHandler RenderActionChanged;

        public event PreviewRequestHandler PreviewRequested;

        public RenderHandler()
        {
            _requestPreviewTimer = new Timer(5000); // every 5s
            _requestPreviewTimer.Elapsed += RequestPreview;
            RenderProgressDialog = new OutputViewerDialog("Output Viewer Dialog");
            RenderProgressDialog.ViewModel.RenderCanceled += CancelRender;
            RenderProgressDialog.ViewModel.RenderCanceledAndSave += CancelAndSaveRender;
        }

        public OutputViewerDialog RenderProgressDialog { get; }

        public Stream BitmapSourceToStream(BitmapSource bitmapSource)
        {
            MemoryStream stream = new MemoryStream();
            BitmapEncoder encoder = new PngBitmapEncoder(); // Or any other suitable encoder like JpegBitmapEncoder
            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
            encoder.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        public void CancelRender()
            => RenderActionChanged?.Invoke(RenderAction_e.Cancel);

        public void CancelAndSaveRender()
            => RenderActionChanged?.Invoke(RenderAction_e.CancelAndSave);

        private void RequestPreview(object sender, ElapsedEventArgs e)
        {
            if (_previewRequestPending)
            {
                return;
            }

            float previewY = 300.0f;
            float previewX = 400.0f;

            if (RenderProgressDialog.ViewModel.AspectRatio >= previewX / previewY)
            {
                previewY = previewX / RenderProgressDialog.ViewModel.AspectRatio;
            }
            else
            {
                previewX = previewY / RenderProgressDialog.ViewModel.AspectRatio;
            }

            PreviewRequested?.Invoke((int)previewX, (int)previewY);
            _previewRequestPending = true;
        }
            

        public void RendererPropertyChangedNotify(RenderProperty_e renderProperty,
            IRendererPropertiesStatus rendererPropertiesStatus)
        {
            if (!_requestPreviewTimer.Enabled)
            {
                _requestPreviewTimer.Start();
            }

            if (renderProperty == RenderProperty_e.CurrentProgress)
            {
                RenderProgressDialog.ViewModel.RenderProgress = (double)(rendererPropertiesStatus.CurrentProgress * 100);
            }
            else if (renderProperty == RenderProperty_e.FramesRendered)
            {
                RenderProgressDialog.ViewModel.FramesRendered = rendererPropertiesStatus.FramesRendered;
            }
            else if (renderProperty == RenderProperty_e.ElapsedTime)
            {
                RenderProgressDialog.ViewModel.TimeSpan = rendererPropertiesStatus.ElapsedTime;
            }
            else if (renderProperty == RenderProperty_e.RenderJobStatus)
            {
                if (rendererPropertiesStatus.RenderJobStatus == RenderJobStatus_e.JobComplete ||
                    rendererPropertiesStatus.RenderJobStatus == RenderJobStatus_e.JobCancelled ||
                    rendererPropertiesStatus.RenderJobStatus == RenderJobStatus_e.JobAborted ||
                    rendererPropertiesStatus.RenderJobStatus == RenderJobStatus_e.JobFailed)
                {
                    _requestPreviewTimer.Stop();
                    RenderProgressDialog.ViewModel.RenderCanceled -= CancelRender;
                    RenderProgressDialog.ViewModel.RenderCanceledAndSave -= CancelAndSaveRender;
                }
            }
            else if (renderProperty == RenderProperty_e.RendererPreviewComplete)
            {
                if (rendererPropertiesStatus.BitmapPreviewImageSource != null)
                {
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.StreamSource = BitmapSourceToStream(rendererPropertiesStatus.BitmapPreviewImageSource);
                    bitmapImage.EndInit();
                    RenderProgressDialog.ViewModel.RenderedImage = bitmapImage;

                    _previewRequestPending = false;
                }
            }
        }
    }
}
