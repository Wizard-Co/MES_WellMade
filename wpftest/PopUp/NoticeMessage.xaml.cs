using System.Windows;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// NoticeMessage.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class NoticeMessage : Window
    {
        public NoticeMessage()
        {
            InitializeComponent();
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Application curApp = Application.Current;
            Window mainWindow = curApp.MainWindow;
            this.Left = mainWindow.Left + (mainWindow.Width - this.ActualWidth) / 2;
            this.Top = mainWindow.Top + (mainWindow.Height - this.ActualHeight) / 2;

        }
    }
}
