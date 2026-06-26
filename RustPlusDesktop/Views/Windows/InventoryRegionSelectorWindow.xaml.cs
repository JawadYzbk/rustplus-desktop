using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace RustPlusDesk.Views
{
    public partial class InventoryRegionSelectorWindow : Window
    {
        private Point _startPoint;
        private bool _isDragging;
        
        public Rect SelectedRegion { get; private set; }

        public InventoryRegionSelectorWindow(BitmapSource screenshot)
        {
            InitializeComponent();
            ImgScreenshot.Source = screenshot;

            // Cover the screen containing primary bounds
            var bounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
            
            // Convert pixels to DIPs for sizing the WPF window
            var (scaleX, scaleY) = GetDpiScale();
            Left = bounds.X / scaleX;
            Top = bounds.Y / scaleY;
            Width = bounds.Width / scaleX;
            Height = bounds.Height / scaleY;

            FullWindowGeometry.Rect = new Rect(0, 0, Width, Height);

            Loaded += (s, e) =>
            {
                Focus();
                CaptureMouse();
            };
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.ChangedButton == MouseButton.Left)
            {
                _startPoint = e.GetPosition(this);
                _isDragging = true;

                SelectionBorder.Width = 0;
                SelectionBorder.Height = 0;
                SelectionBorder.Visibility = Visibility.Visible;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_isDragging)
            {
                Point currentPoint = e.GetPosition(this);

                double x = Math.Min(_startPoint.X, currentPoint.X);
                double y = Math.Min(_startPoint.Y, currentPoint.Y);
                double w = Math.Max(0.1, Math.Abs(_startPoint.X - currentPoint.X));
                double h = Math.Max(0.1, Math.Abs(_startPoint.Y - currentPoint.Y));

                Rect selectedRect = new Rect(x, y, w, h);
                SelectionGeometry.Rect = selectedRect;

                Canvas.SetLeft(SelectionBorder, x);
                Canvas.SetTop(SelectionBorder, y);
                SelectionBorder.Width = w;
                SelectionBorder.Height = h;
            }
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.ChangedButton == MouseButton.Left && _isDragging)
            {
                _isDragging = false;
                ReleaseMouseCapture();

                Point currentPoint = e.GetPosition(this);
                double x = Math.Min(_startPoint.X, currentPoint.X);
                double y = Math.Min(_startPoint.Y, currentPoint.Y);
                double w = Math.Max(0.1, Math.Abs(_startPoint.X - currentPoint.X));
                double h = Math.Max(0.1, Math.Abs(_startPoint.Y - currentPoint.Y));

                SelectedRegion = new Rect(x, y, w, h);

                if (w > 10 && h > 10)
                {
                    DialogResult = true;
                    Close();
                }
                else
                {
                    SelectionBorder.Visibility = Visibility.Collapsed;
                    SelectionGeometry.Rect = Rect.Empty;
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private static (double scaleX, double scaleY) GetDpiScale()
        {
            try
            {
                var mainWin = Application.Current.MainWindow;
                if (mainWin != null)
                {
                    var source = PresentationSource.FromVisual(mainWin);
                    if (source?.CompositionTarget != null)
                    {
                        return (source.CompositionTarget.TransformToDevice.M11,
                                source.CompositionTarget.TransformToDevice.M22);
                    }
                }
            }
            catch { }

            using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                return (g.DpiX / 96.0, g.DpiY / 96.0);
            }
        }
    }
}
