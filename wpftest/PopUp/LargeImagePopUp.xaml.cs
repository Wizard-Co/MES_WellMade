using System.Windows;
using System.Windows.Media.Imaging;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// LargeImagePopUp.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class LargeImagePopUp : Window
    {
        BitmapImage TheImage { get; set; }

        public LargeImagePopUp()
        {
            InitializeComponent();
        }

        public LargeImagePopUp(BitmapImage bitmapImage)
        {
            //TheImage = bitmapImage;
            InitializeComponent();
            LargeImage.Source = bitmapImage;
        }
    }
}
