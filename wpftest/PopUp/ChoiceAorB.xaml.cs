using System.Windows;
using System.Windows.Controls;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// ChoiceAorB.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ChoiceAorB : Window
    {
        public string Choice = string.Empty;

        public ChoiceAorB()
        {
            InitializeComponent();
        }

        private void ChoiceAorB_Loaded(object sender, RoutedEventArgs e)
        {
            Application curApp = Application.Current;
            Window mainWindow = curApp.MainWindow;
            this.Left = mainWindow.Left + (mainWindow.Width - this.ActualWidth) / 2;
            this.Top = mainWindow.Top + (mainWindow.Height - this.ActualHeight) / 2;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Choice_Click(object sender, RoutedEventArgs e)
        {
            if (((Button)sender).Name == "A_Button")
            {
                Choice = "A";
            }
            else
            {
                Choice = "B";
            }
            DialogResult = DialogResult.HasValue;
        }
    }
}
