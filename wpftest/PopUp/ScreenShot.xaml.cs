using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// ScreenShot.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ScreenShot : Window
    {
        private Point _origin; // Original Offset of image
        private Point _start; // Original Position of the mouse

        public ScreenShot()
        {
            InitializeComponent();
            this.KeyDown += ScreenShot_KeyDown;
        }

        private void ScreenShot_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.W && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                this.Close();
                e.Handled = true;
            }
            else if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                btnSaveNameOther_Click(null, null);
                e.Handled = true;
            }
            //else if (e.Key == Key.D0 && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            //{
            //    // Ctrl+0: 100% 크기
            //    btnZoom100_Click(null, null);
            //    e.Handled = true;
            //}
            else if (e.Key == Key.D1 && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                // Ctrl+1: 창에 맞춤
                btnZoomFit_Click(null, null);
                e.Handled = true;
            }
            else if (e.Key == Key.Add && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                // Ctrl+Plus: 확대
                ZoomIn();
                e.Handled = true;
            }
            else if (e.Key == Key.Subtract && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                // Ctrl+Minus: 축소
                ZoomOut();
                e.Handled = true;
            }
        }

        private void ScreenShot_Loaded(object sender, RoutedEventArgs e)
        {
            if (MainWindow.ScreenCapture != null && MainWindow.ScreenCapture.Count > 0)
            {
                ImageData.Source = MainWindow.ScreenCapture[0].Source;

                // 기존처럼 Stretch="Uniform"으로 시작 시 창에 맞게 표시
                // 줌 리셋으로 원본 크기로 변경 가능
            }
        }

        // 마우스 휠 이벤트로 확대/축소
        private void ScreenShot_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (ImageData.Source == null) return;

            // 마우스 위치를 이미지 기준으로 가져오기
            Point mousePos = e.GetPosition(ImageData);
            Matrix matrix = ImageData.RenderTransform.Value;

            if (e.Delta > 0)
            {
                // 확대 (1.1배)
                matrix.ScaleAtPrepend(1.1, 1.1, mousePos.X, mousePos.Y);
            }
            else
            {
                // 축소 (1/1.1배)
                matrix.ScaleAtPrepend(1.0 / 1.1, 1.0 / 1.1, mousePos.X, mousePos.Y);
            }

            ImageData.RenderTransform = new MatrixTransform(matrix);
            e.Handled = true;

            // 타이틀에 줌 레벨 표시
            double zoomLevel = matrix.M11 * 100; // M11이 X축 스케일
            this.Title = $"이미지 보기 - {zoomLevel:F0}%";
        }

        // 마우스 드래그 시작
        private void ImageData_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ImageData.IsMouseCaptured) return;

            ImageData.CaptureMouse();
            _start = e.GetPosition(ImageBorder);
            _origin.X = ImageData.RenderTransform.Value.OffsetX;
            _origin.Y = ImageData.RenderTransform.Value.OffsetY;

            ImageData.Cursor = Cursors.Hand;
        }

        // 마우스 드래그 중
        private void ImageData_MouseMove(object sender, MouseEventArgs e)
        {
            if (!ImageData.IsMouseCaptured) return;

            Point currentPos = e.GetPosition(ImageBorder);
            Matrix matrix = ImageData.RenderTransform.Value;

            matrix.OffsetX = _origin.X + (currentPos.X - _start.X);
            matrix.OffsetY = _origin.Y + (currentPos.Y - _start.Y);

            ImageData.RenderTransform = new MatrixTransform(matrix);
        }

        // 마우스 드래그 종료
        private void ImageData_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ImageData.ReleaseMouseCapture();
            ImageData.Cursor = Cursors.Arrow;
        }

        //우클릭 메뉴 닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // 100% 크기 (현재 창에 맞춘 상태에서 100% 스케일)
        private void btnZoom100_Click(object sender, RoutedEventArgs e)
        {
            if (ImageData.Source == null) return;

            // Stretch는 그대로 유지하고 Transform만 1.0 스케일로 설정
            Matrix matrix = new Matrix();
            matrix.ScaleAt(1.0, 1.0, ImageData.ActualWidth / 2, ImageData.ActualHeight / 2);

            ImageData.RenderTransform = new MatrixTransform(matrix);
            this.Title = "이미지 보기 - 100%";
        }

        // 창에 맞춤
        private void btnZoomFit_Click(object sender, RoutedEventArgs e)
        {
            // Stretch를 Uniform으로 변경하고 Transform 초기화
            ImageData.Stretch = Stretch.Uniform;
            ImageData.RenderTransform = new MatrixTransform();
            this.Title = "이미지 보기 - 창에 맞춤";
        }

        // 중앙 기준 확대
        private void ZoomIn()
        {
            if (ImageData.Source == null) return;

            Point centerPoint = new Point(ImageData.ActualWidth / 2, ImageData.ActualHeight / 2);
            Matrix matrix = ImageData.RenderTransform.Value;
            matrix.ScaleAtPrepend(1.25, 1.25, centerPoint.X, centerPoint.Y);
            ImageData.RenderTransform = new MatrixTransform(matrix);

            // 타이틀 업데이트
            double zoomLevel = matrix.M11 * 100;
            this.Title = $"이미지 보기 - {zoomLevel:F0}%";
        }

        // 중앙 기준 축소
        private void ZoomOut()
        {
            if (ImageData.Source == null) return;

            Point centerPoint = new Point(ImageData.ActualWidth / 2, ImageData.ActualHeight / 2);
            Matrix matrix = ImageData.RenderTransform.Value;
            matrix.ScaleAtPrepend(1.0 / 1.25, 1.0 / 1.25, centerPoint.X, centerPoint.Y);
            ImageData.RenderTransform = new MatrixTransform(matrix);

            // 타이틀 업데이트
            double zoomLevel = matrix.M11 * 100;
            this.Title = $"이미지 보기 - {zoomLevel:F0}%";
        }

        //우클릭 메뉴 다른이름으로 저장
        private void btnSaveNameOther_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ImageData.Source == null)
                {
                    MessageBox.Show("저장할 이미지가 없습니다.", "확인");
                    return;
                }

                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "PNG 파일|*.png|JPG 파일|*.jpg|BMP 파일|*.bmp",
                    FileName = "Image" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss")
                };

                if (saveDialog.ShowDialog() == true)
                {
                    BitmapSource bitmap = ImageData.Source as BitmapSource;

                    // 기본 WPF배경색 #F0F0F0 색 채우기
                    var visual = new DrawingVisual();
                    using (var context = visual.RenderOpen())
                    {
                        context.DrawRectangle(new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0)), null,
                            new Rect(0, 0, bitmap.PixelWidth, bitmap.PixelHeight));
                        context.DrawImage(bitmap, new Rect(0, 0, bitmap.PixelWidth, bitmap.PixelHeight));
                    }

                    // 컨트롤들 채우기
                    var renderTarget = new RenderTargetBitmap(bitmap.PixelWidth, bitmap.PixelHeight, 96, 96, PixelFormats.Pbgra32);
                    renderTarget.Render(visual);

                    // 파일 확장자에 따른 인코더 선택
                    BitmapEncoder encoder;
                    string extension = System.IO.Path.GetExtension(saveDialog.FileName).ToLower();

                    switch (extension)
                    {
                        case ".jpg":
                        case ".jpeg":
                            encoder = new JpegBitmapEncoder() { QualityLevel = 95 };
                            break;
                        case ".bmp":
                            encoder = new BmpBitmapEncoder();
                            break;
                        default:
                            encoder = new PngBitmapEncoder();
                            break;
                    }

                    encoder.Frames.Add(BitmapFrame.Create(renderTarget));

                    using (FileStream stream = new FileStream(saveDialog.FileName, FileMode.Create))
                    {
                        encoder.Save(stream);
                    }

                    MessageBox.Show($"저장 되었습니다.\n경로: {saveDialog.FileName}", "확인");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 오류: {ex.Message}", "확인");
            }
        }
    }
}