using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// MultiLevelSort.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MultiLevelSort : Window
    {
        DataGrid TargetDG;
        Lib lib = new Lib();

        public List<string> ColName = new List<string>();
        public List<string> SortingProperty = new List<string>();
        public List<int> ColIndex = new List<int>();


        public MultiLevelSort(DataGrid TargetDatagrid)
        {
            TargetDG = TargetDatagrid;
            InitializeComponent();
        }


        private void MultiLevelSort_Loaded(object sender, RoutedEventArgs e)
        {
            var data = new SortingData { RankingNUM = "1순위", ColName = "", SortingProperty = "" };
            dgdSorting.Items.Add(data);
        }


        // 콤보박스 정보 가져오기.  (열 명 콤보박스)
        private void ComboBoxSetting(object sender, RoutedEventArgs e)
        {
            ComboBox cbodgdColName = (ComboBox)sender;
            cbodgdColName.ItemsSource = null;

            List<string> lis = new List<string>();

            for (int i = 0; i < TargetDG.Columns.Count; i++)
            {
                //if (TargetDG.Columns[i].GetType() == typeof(DataGridTextColumn))
                //{
                lis.Add(TargetDG.Columns[i].Header.ToString());
                //}
            }

            cbodgdColName.ItemsSource = lis;
            cbodgdColName.SelectedIndex = 0;

        }


        // 콤보박스 정보 가져오기.  (정렬 콤보박스)
        private void ComboBoxSetting2(object sender, RoutedEventArgs e)
        {
            ComboBox cbodgdSortingProperty = (ComboBox)sender;
            cbodgdSortingProperty.ItemsSource = null;

            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            DataRow row0 = dt.NewRow();
            row0["value"] = "UP";
            row0["display"] = "오름차순";

            DataRow row1 = dt.NewRow();
            row1["value"] = "DOWN";
            row1["display"] = "내림차순";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);

            cbodgdSortingProperty.ItemsSource = dt.DefaultView;
            cbodgdSortingProperty.DisplayMemberPath = "display";
            cbodgdSortingProperty.SelectedValuePath = "value";
            cbodgdSortingProperty.SelectedIndex = 0;
        }


        // 행 추가.
        private void btnRowAdd_Click(object sender, RoutedEventArgs e)
        {
            int nowCount = dgdSorting.Items.Count + 1;
            string nCount = nowCount.ToString();

            var data = new SortingData { RankingNUM = nCount + "순위", ColName = "", SortingProperty = "" };
            dgdSorting.Items.Add(data);
        }
        // 행 삭제.
        private void btnRowDelete_Click(object sender, RoutedEventArgs e)
        {
            //1. 일반적인 빼기과정.
            var ViewReceiver = dgdSorting.CurrentItem as SortingData;  //선택 줄.
            if (ViewReceiver == null)   // 선택한 줄이 없다면,
            {
                int Point = dgdSorting.Items.Count;
                DataGridRow dgr = lib.GetRow(Point - 1, dgdSorting);
                ViewReceiver = dgr.Item as SortingData;
            }

            dgdSorting.Items.Remove(ViewReceiver);
        }


        // 확인버튼 클릭 시.
        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            int UpgradePoint = dgdSorting.Items.Count;
            for (int i = 0; i < UpgradePoint; i++)
            {
                DataGridRow dgr = lib.GetRow(i, dgdSorting);

                DataGridCell cell0 = lib.GetCell(i, 1, dgdSorting);
                ComboBox cb0 = lib.GetVisualChild<ComboBox>(cell0);
                DataGridCell cell1 = lib.GetCell(i, 2, dgdSorting);
                ComboBox cb1 = lib.GetVisualChild<ComboBox>(cell1);

                if (cb0.SelectedValue != null)
                {
                    ColName.Add(cb0.SelectedValue.ToString());

                    for (int j = 0; j < TargetDG.Columns.Count; j++)
                    {
                        if (TargetDG.Columns[j].Header.ToString() == cb0.SelectedValue.ToString())
                        {
                            ColIndex.Add(j);
                        }
                    }
                }
                if (cb1.SelectedValue != null)
                {
                    SortingProperty.Add(cb1.SelectedValue.ToString());
                }
            }
            DialogResult = DialogResult.HasValue;
        }


        // 취소 버튼 클릭시.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


    }


    public class SortingData
    {
        public string RankingNUM { get; set; }
        public string ColName { get; set; }
        public string SortingProperty { get; set; }
    }
}
