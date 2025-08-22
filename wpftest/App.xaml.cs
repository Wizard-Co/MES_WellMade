using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace WizMes_WellMade
{
    /// <summary>
    /// App.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class App : Application
    {
        Lib lib = new Lib();

        public App()
        {
            //mfont = "fonts/#궁서체";
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(ResolveAssembly);
            LoadINI loini = new LoadINI();
            loini.loadINI();
        }

        static Assembly ResolveAssembly(object sender, ResolveEventArgs args)
        {
            //We dont' care about System Assembies and so on...
            //if (!args.Name.ToLower().StartsWith("Test")) return null;

            Assembly thisAssembly = Assembly.GetExecutingAssembly();

            //Get the Name of the AssemblyFile
            var name = args.Name.Substring(0, args.Name.IndexOf(',')) + ".dll";

            //Load form Embedded Resources - This Function is not called if the Assembly is in the Application Folder
            var resources = thisAssembly.GetManifestResourceNames().Where(s => s.EndsWith(name));
            if (resources.Count() > 0)
            {
                var resourceName = resources.First();
                using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return null;
                    var block = new byte[stream.Length];
                    stream.Read(block, 0, block.Length);
                    return Assembly.Load(block);
                }
            }
            return null;
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Create the startup window
            MainWindow wnd = new MainWindow();
            // Do stuff here, e.g. to the window
            // Show the window
            wnd.Show();
        }

        //
        private void TextBoxZero_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringZero(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        //
        private void TextBoxOne_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringOne(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        //
        private void TextBoxTwo_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringTwo(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        private void TextBoxZero_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringZero(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        private void TextBoxOne_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringTwoExceptDot(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        private void TextBoxTwo_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;

            if (tb != null)
            {
                if (Lib.Instance.IsNumOrAnother(tb.Text))
                {
                    tb.Text = Lib.Instance.returnNumStringTwoExceptDot(tb.Text);
                    tb.SelectionStart = tb.Text.Length;
                    sender = tb;
                }
            }
        }

        private void MouseLeftDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                try
                {
                    UserControl userControl = Lib.Instance.GetParent<UserControl>(sender as DataGrid);
                    if (userControl != null)
                    {
                        object objUpdate = userControl.FindName("btnUpdate");
                        object objEdit = userControl.FindName("btnEdit");

                        if (objUpdate != null)
                        {
                            if ((objUpdate as Button).IsEnabled == true)
                            {
                                (objUpdate as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                            }
                        }
                        else if (objEdit != null)
                        {
                            if ((objEdit as Button).IsEnabled == true)
                            {
                                (objEdit as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        // 데이트피커에 값(숫자)을 입력후에 enter 쳤을때!!!!!!!!! 그걸 날짜로 적용해서 해당 데이트피커에 적용되도록!!!!
        private void DatePicker_EnterDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DatePicker dtpSender = sender as DatePicker;

                DateTime dt = new DateTime();
                if (DateTime.TryParse(DatePickerFormat(dtpSender.Text), out dt))
                {
                    dtpSender.SelectedDate = dt;
                }
            }
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            str = str.Trim().Replace("-", "").Replace(".", "");

            if (!str.Equals(""))
            {
                if (str.Length == 8)
                {
                    str = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
                else if (str.Length == 7)
                {
                    str = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-0" + str.Substring(6, 1);
                }
                else if (str.Length == 6)
                {
                    str = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-01";
                }
                else if (str.Length == 5)
                {
                    str = str.Substring(0, 4) + "-0" + str.Substring(4, 1) + "-01";
                }
                else if (str.Length == 4)
                {
                    str = DateTime.Today.ToString("yyyy") + "-" + str.Substring(0, 2) + "-" + str.Substring(2, 2);
                }
            }

            return str;
        }

        private void TextBox_Loaded(object sender, RoutedEventArgs e)
        {
            TextBox txtSender = sender as TextBox;

            string name = txtSender.Name;

            if (txtSender.TextWrapping != TextWrapping.Wrap)
            {
                double txtHeight = txtSender.FontSize;

                if (txtSender.ActualHeight != 0
                    && txtSender.ActualHeight < txtHeight)
                {
                    double fontSize = txtSender.ActualHeight / 2;

                    txtSender.FontSize = fontSize;
                }
            }
        }
        private void TextBox_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                TextBox txtSender = sender as TextBox;

                Size n = e.NewSize;
                Size p = e.PreviousSize;
                double l = n.Height / p.Height;
                if (l != double.PositiveInfinity
                    && l != 0
                    && double.IsNaN(l) == false)
                {
                    txtSender.FontSize = txtSender.FontSize * l;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("에러 : " + ex.Message);
            }
        }

        private void DatePicker_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                DatePicker dtpSender = sender as DatePicker;

                int length = dtpSender.Text.Length == 0 ? 10 : dtpSender.Text.Length;
                double txtWidth = dtpSender.FontSize * length;

                // 폭이 맞지 않다면.
                if (dtpSender.ActualWidth != 0
                    && dtpSender.ActualWidth < txtWidth)
                {
                    double result = dtpSender.ActualWidth / txtWidth;
                    dtpSender.FontSize = dtpSender.FontSize * result;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void DatePicker_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                DatePicker dtpSender = sender as DatePicker;

                Size n = e.NewSize;
                Size p = e.PreviousSize;

                if (n.Width != p.Width)
                {
                    double l = n.Width / p.Width;
                    if (l != double.PositiveInfinity
                        && l != 0
                        && double.IsNaN(l) == false)
                    {
                        double txtWidth = dtpSender.Text.Length * dtpSender.FontSize;

                        double maxWidth = dtpSender.Text.Length * WizMes_WellMade.MainWindow.StdFontSize;

                        // 줄어들 때 : 
                        if (n.Width < p.Width)
                        {
                            if (n.Width < txtWidth)
                            {
                                Console.Write(n.Width);
                                Console.Write(txtWidth);

                                dtpSender.FontSize = dtpSender.FontSize * l;
                            }
                        }
                        // 늘어날 때 : 
                        else if (n.Width > p.Width)
                        {
                            // Width 값으로 계산을 하려니.. MainWindow.Fontsiz
                            double maxFontSize = WizMes_WellMade.MainWindow.StdFontSize;

                            double changeFontSize = dtpSender.FontSize * l;

                            if (maxFontSize > changeFontSize)
                            {
                                dtpSender.FontSize = dtpSender.FontSize * l;
                            }

                            //if (maxWidth > n.Width)
                            //{
                            //    dtpSender.FontSize = dtpSender.FontSize * l;
                            //}
                        }
                    }
                }
                else if (n.Height != p.Height)
                {
                    double l = n.Height / p.Height;
                    if (l != double.PositiveInfinity
                        && l != 0
                        && double.IsNaN(l) == false)
                    {
                        dtpSender.FontSize = dtpSender.FontSize * l;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("에러 : " + ex.Message);
            }
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                ComboBox cboSender = sender as ComboBox;

                double txtHeight = cboSender.FontSize;

                if (cboSender.ActualHeight != 0
                    && cboSender.ActualHeight < txtHeight)
                {
                    double fontSize = cboSender.ActualHeight / 2;

                    cboSender.FontSize = fontSize;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ComboBox_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                ComboBox cboSender = sender as ComboBox;

                Size n = e.NewSize;
                Size p = e.PreviousSize;
                if (n.Width != p.Width)
                {
                    double l = n.Width / p.Width;
                    if (l != double.PositiveInfinity
                        && l != 0
                        && double.IsNaN(l) == false)
                    {
                        double txtWidth = (cboSender.Text.Length + 2) * cboSender.FontSize;

                        double maxWidth = cboSender.Text.Length * WizMes_WellMade.MainWindow.StdFontSize;

                        // 줄어들 때 : 
                        if (n.Width < p.Width)
                        {
                            if (n.Width < txtWidth)
                            {
                                Console.Write(n.Width);
                                Console.Write(txtWidth);

                                cboSender.FontSize = cboSender.FontSize * l;
                            }
                        }
                        // 늘어날 때 : 
                        else if (n.Width > p.Width)
                        {
                            // Width 값으로 계산을 하려니.. MainWindow.Fontsiz
                            double maxFontSize = WizMes_WellMade.MainWindow.StdFontSize;

                            double changeFontSize = cboSender.FontSize * l;

                            if (maxFontSize > changeFontSize)
                            {
                                cboSender.FontSize = cboSender.FontSize * l;
                            }

                            //if (maxWidth > n.Width)
                            //{
                            //    dtpSender.FontSize = dtpSender.FontSize * l;
                            //}
                        }
                    }
                }
                else if (n.Height != p.Height)
                {
                    double l = n.Height / p.Height;
                    if (l != double.PositiveInfinity
                        && l != 0
                        && double.IsNaN(l) == false)
                    {
                        cboSender.FontSize = cboSender.FontSize * l;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("에러 : " + ex.Message);
            }
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;

            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }

            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;


            //System.Diagnostics.Debug.WriteLine("값1- " + a);
            //System.Diagnostics.Debug.WriteLine("값2- " + b);
            //System.Diagnostics.Debug.WriteLine("값3- " + c);

            //System.Diagnostics.Debug.WriteLine("값4- " + dgs.ColumnHeaderHeight);
            //System.Diagnostics.Debug.WriteLine("값4- " + dgs.FontSize);

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.RowHeight = dgs.RowHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }

        private void TotalDataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;

            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }

            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;


            //System.Diagnostics.Debug.WriteLine("값1- " + a);
            //System.Diagnostics.Debug.WriteLine("값2- " + b);
            //System.Diagnostics.Debug.WriteLine("값3- " + c);

            //System.Diagnostics.Debug.WriteLine("값4- " + dgs.ColumnHeaderHeight);
            //System.Diagnostics.Debug.WriteLine("값4- " + dgs.FontSize);

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                //dgs.RowHeight = dgs.RowHeight * c;
                //dgs.FontSize = dgs.FontSize * c;
            }
        }

        private void TextBoxOnlyNumber_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        private void TextBoxOnlyNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key.Equals(Key.ImeProcessed))
            {
                e.Handled = true;
            }
        }

        private void TextBoxOnlyNumber_Integer_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumericOnly((TextBox)sender, e);
        }
    }
}
