using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// ScreenShot.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ScreenShot : Window
    {
        public ScreenShot()
        {
            InitializeComponent();
        }

        private void ScreenShot_Loaded(object sender, RoutedEventArgs e)
        {
            if (MainWindow.ScreenCapture != null && MainWindow.ScreenCapture.Count > 0)
            {
                ImageData.Source = MainWindow.ScreenCapture[0].Source;
            }
        }

        //우클릭 메뉴 닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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
                    Filter = "PNG 파일|*.png",
                    FileName = "Screenshot_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss")
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

                    //인코딩 및 저장
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderTarget));
                    using (FileStream stream = new FileStream(saveDialog.FileName, FileMode.Create))
                    {
                        encoder.Save(stream);
                    }
                    MessageBox.Show("저장 되었습니다.", "확인");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 오류: {ex.Message}", "확인");
            }
        }
    }
}
