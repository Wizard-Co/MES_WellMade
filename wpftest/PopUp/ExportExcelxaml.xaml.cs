using System.Windows;

namespace WizMes_WellMade.PopUP
{
    /// <summary>
    /// ExportExcelxaml.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ExportExcelxaml : Window
    {
        string[] Temp = null;
        public string choice = string.Empty;
        public string Check = string.Empty;

        public ExportExcelxaml()
        {
            InitializeComponent();
        }

        public ExportExcelxaml(string[] list)
        {
            InitializeComponent();
            Temp = list;
        }

        public void SetGrid(string[] list)
        {
            for (int i = 0; i < list.Length / 2; i++)
            {
                var data = new ExcelData { Name = list[i], PropertyName = list[i + list.Length / 2] };
                excelName.Items.Add(data);
            }
        }

        private void ExportExcel_Loaded(object sender, RoutedEventArgs e)
        {
            Application curApp = Application.Current;
            Window mainWindow = curApp.MainWindow;
            this.Left = mainWindow.Left + (mainWindow.Width - this.ActualWidth) / 2;
            this.Top = mainWindow.Top + (mainWindow.Height - this.ActualHeight) / 2;
            SetGrid(Temp);
            excelName.SelectedIndex = 0;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelData ecd = excelName.SelectedItem as ExcelData;
            if (ecd != null)
            {
                choice = ecd.PropertyName;

                if (FormCheck.IsChecked == true)
                {
                    Check = "Y";
                }
                else
                {
                    Check = "N";
                }
            }

            DialogResult = DialogResult.HasValue;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class ExcelData
    {
        public string Name { get; set; }
        public string PropertyName { get; set; }
    }
}
