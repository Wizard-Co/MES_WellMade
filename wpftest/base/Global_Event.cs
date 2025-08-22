using System.Windows.Controls;
using System.Windows.Input;

namespace WizMes_WellMade
{
    public static class Global_Event
    {
        public static RoutedCommand ClickCommand { get; }
        public static TextChangedEventHandler DataGridInHeaderChanged { get; }

        /// <summary>
        /// filter 할 table
        /// </summary>
        private static DataGrid FilterDataGrid { get; set; }
        /// <summary>
        /// filter 전 table
        /// </summary>
        private static DataGrid FilterDataGridPre { get; set; }
        /// <summary>
        /// filter 후 table
        /// </summary>
        private static DataGrid FilterDataGridAfter { get; set; }

        static Global_Event()
        {
            ClickCommand = new RoutedCommand("ClickCommand", typeof(Global_Event));
            CommandManager.RegisterClassCommandBinding(
              typeof(Button),
              new CommandBinding(ClickCommand, OnColumnHeaderClick));
            DataGridInHeaderChanged = new TextChangedEventHandler(OnDataGridInHeaderTextChanged);
        }

        static void OnColumnHeaderClick(object sender, ExecutedRoutedEventArgs e)
        {

        }

        static void OnDataGridInHeaderTextChanged(object sender, TextChangedEventArgs e)
        {
            //MdiChild mdiChild = MainWindow.MainMdiContainer.ActiveMdiChild;
            //UserControl CurrentControl = mdiChild.Content as UserControl;

            //DataTable BaseDataTable = new DataTable();
            //DataTable WhereDataTable = new DataTable();

            //TextBox textBox = sender as TextBox;
            //string plusSql = " where ";

            //DataTable FilterText(string strFilter)
            //{
            //    DataTable returnTable = null;
            //    DataGridColumnHeader dgch = Lib.Instance.GetParent<DataGridColumnHeader>(sender as TextBox);

            //    if (dgch != null)
            //    {
            //        if (Lib.Instance.IsNumOrAnother(dgch.Tag.ToString()))
            //        {
            //            int colNum = Convert.ToInt32(dgch.Tag);
            //            BaseDataTable = Lib.Instance.DataGirdToDataTable(FilterDataGridPre);
            //            returnTable = BaseDataTable.Clone();

            //            if (!strFilter.Equals(string.Empty))
            //            {
            //                plusSql += BaseDataTable.Columns[colNum].Caption+" LIKE  '%'"+strFilter+"'%' ";

            //                foreach (DataRow dr in BaseDataTable.Select(plusSql))
            //                {
            //                    returnTable.Rows.Add(dr.ItemArray);
            //                }
            //            }
            //        }
            //    }

            //    return returnTable;
            //}

            //if (CurrentControl != null)
            //{
            //    object data = CurrentControl.Tag;
            //    FilterDataGridPre = data as DataGrid;

            //    if (data != null)
            //    {
            //        WhereDataTable = FilterText(textBox.Text);

            //        if (WhereDataTable != null && WhereDataTable.Rows.Count > 0)
            //        {
            //            foreach (DataRow dr in WhereDataTable.Rows)
            //            {
            //                FilterDataGridAfter.Items.Add(dr.ItemArray);
            //            }
            //        }
            //    }
            //}
        }
    }
}
